# -*- coding: utf-8 -*-
"""
审查模块 - 调用豆包（火山引擎）API 逐条审查
"""
import json
import os
import base64
import io
import re
import time
from PIL import Image
from openai import OpenAI
import httpx
import fitz
import config
from file_index import _terms


class Reviewer:
    """项目规范性审查器"""

    def __init__(self, api_key=None):
        # 豆包客户端（同时支持文本和视觉审查）
        self.api_key = api_key or config.VOLC_API_KEY
        self.base_url = config.VOLC_BASE_URL
        # 某些本机环境会注入不可用的 127.0.0.1:9 代理，导致请求尚未到达 Ark。
        # 直连火山接口，Web 页面仍可通过 api_key 临时覆盖本地配置。
        http_client = httpx.Client(
            trust_env=False,
            timeout=httpx.Timeout(config.API_REQUEST_TIMEOUT_SECONDS, connect=20.0),
        )
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
            http_client=http_client,
        )

    def _extract_response_text(self, response):
        """兼容多种 API 返回格式提取文本"""
        # Responses API 格式
        if hasattr(response, 'output') and response.output:
            try:
                return response.output[0].content[0].text
            except (AttributeError, IndexError, TypeError):
                pass

        # Chat Completions 格式
        if hasattr(response, 'choices') and response.choices:
            try:
                return response.choices[0].message.content
            except (AttributeError, IndexError, TypeError):
                pass

        # 直接返回字符串
        if isinstance(response, str):
            return response

        # 兜底：转为字符串
        return str(response)

    def _create_completion(self, **kwargs):
        """网络短暂波动时重试，避免把连接错误直接写成审查结论。"""
        last_error = None
        for attempt in range(config.API_REQUEST_RETRIES):
            try:
                return self.client.chat.completions.create(**kwargs)
            except Exception as exc:
                last_error = exc
                if attempt < config.API_REQUEST_RETRIES - 1:
                    time.sleep(1.5 * (attempt + 1))
        raise last_error

    def review_single_checkpoint(self, checkpoint, file_texts, file_signatures, project_name='', matched_file=''):
        """
        审查单个检查要点（纯文本，不发送图片）
        """
        prompt = self._build_single_prompt(checkpoint, file_texts, file_signatures, project_name, matched_file)

        try:
            response = self._create_completion(
                model=config.VOLC_VISION_MODEL,
                messages=[
                    {"role": "system", "content": self._get_system_prompt()},
                    {"role": "user", "content": [{"type": "text", "text": prompt}]}
                ],
                temperature=0.1,
                max_tokens=2000,
            )
            result_text = self._extract_response_text(response)
            return self._parse_review_result(result_text, checkpoint)
        except Exception as e:
            return {
                'seq': checkpoint['seq'],
                'check_point': checkpoint['check_point'],
                'doc_name': checkpoint['doc_name'],
                'stage': checkpoint['stage'],
                'link': checkpoint['link'],
                'result': '无法判断',
                'issue': f'审查过程出错: {str(e)}',
                'suggestion': '',
                'evidence': '',
                'matched_file': matched_file,
            }

    def review_with_vision(self, checkpoint, file_texts, file_signatures, project_name='', vision_model=None, matched_file=''):
        """
        使用视觉能力审查（发送页面图片给豆包视觉模型）
        """
        # 用户未选择视觉模型 或 未配置 API Key 时，降级为纯文本审查
        if vision_model != 'doubao' or not self.api_key or 'your-volc' in self.api_key.lower():
            return self.review_single_checkpoint(checkpoint, file_texts, file_signatures, project_name, matched_file)

        # 从 file_infos 中提取带文件名的图片分组
        # 优先使用 page_images(bytes)，如果为空则尝试从 page_image_paths(磁盘路径) 加载
        scanned_files = []
        # 视觉识别的成本远高于文本检索。匹配阶段已选出主证据，因此只传主匹配
        # 文件的页面；其他候选仍以文本片段参与交叉核验，避免每条都重复上传整批扫描件。
        candidate_order = []
        if matched_file:
            candidate_order.append(matched_file)
        for name in checkpoint.get('candidate_files', []):
            if name not in candidate_order:
                candidate_order.append(name)
        # 普通检查只发主证据以控制耗时；签章、日期、分包和系统留痕检查必须把
        # 证据组中的多个文件一起送入视觉模型，否则主文件未命中时会漏读交叉佐证。
        vision_limit = config.MAX_VISION_FILES_PER_CHECKPOINT
        cross_file_keywords = (
            '签字', '盖章', '签章', '日期', '分包', 'IRS', '计划', '批复', '评审',
            '金额', '中标', '乙方', '供应商', '投标', '技术规范', '合同流转',
        )
        if any(keyword in checkpoint.get('check_point', '') for keyword in cross_file_keywords):
            # 交叉事项至少查看五份候选；候选由本地规则预先压缩为最小证据集，
            # 不能再只给模型一份扫描件而要求它完成多文件比对。
            vision_limit = max(vision_limit, 5)
        allowed_files = set(candidate_order[:vision_limit])
        for file_info in file_texts:
            if allowed_files and file_info.get('file_key', file_info.get('filename', '')) not in allowed_files:
                continue
            page_imgs = file_info.get('page_images', [])
            page_numbers = []

            # 扫描版合同、中标通知书等不做全文OCR。对长文件优先传封面、主体条款和末页，
            # 并保留页码；避免58页合同与其他扫描件同时发送后，模型无法聚焦金额/签署信息。
            if file_info.get('is_scanned') and file_info.get('extension') == '.pdf':
                try:
                    pdf = fitz.open(file_info.get('filepath', ''))
                    page_count = len(pdf)
                    filename = file_info.get('filename', '')
                    # 合同核验不能只看封面和末页：金额、工期、开工约定常位于正文中部，
                    # 签署页通常在后部。固定传送前部、中部、后部并保留实际页码。
                    head_pages = 17 if '合同' in filename else 12
                    indexes = list(range(min(head_pages, page_count)))
                    if '合同' in filename and page_count > head_pages + 8:
                        middle = page_count // 2
                        indexes.extend(range(max(0, middle - 2), min(page_count, middle + 3)))
                    indexes.extend(range(max(0, page_count - 4), page_count))
                    indexes = sorted(set(indexes))
                    page_imgs = [
                        pdf[index].get_pixmap(dpi=220 if '合同' in filename else 200).tobytes('jpg', jpg_quality=88)
                        for index in indexes
                    ]
                    page_numbers = [index + 1 for index in indexes]
                    pdf.close()
                except Exception:
                    page_imgs = []

            if not page_imgs:
                # 尝试从磁盘加载（agent_core 保存的图片文件）
                img_paths = file_info.get('page_image_paths', [])
                if img_paths:
                    # 扫描件保留全部页面；文字型PDF只允许发送首页、第二页、倒数第二页和末页。
                    if not file_info.get('is_scanned', False) and len(img_paths) > 4:
                        img_paths = img_paths[:2] + img_paths[-2:]
                    loaded = []
                    for p in img_paths:
                        try:
                            with open(p, 'rb') as f:
                                loaded.append(f.read())
                        except Exception:
                            pass
                    # 取最高品质的 JPEG 格式加载（压缩更小、传输更快）
                    page_imgs = []
                    for raw in loaded:
                        try:
                            img = Image.open(io.BytesIO(raw))
                            buf = io.BytesIO()
                            img.save(buf, format='JPEG', quality=95)
                            page_imgs.append(buf.getvalue())
                        except Exception:
                            page_imgs.append(raw)  # 降级用原始数据
                    # agent_core 的图片文件名保留原 PDF 页码，避免末页被标为第3/4页。
                    page_numbers = []
                    for path in img_paths:
                        found = re.search(r'_page_(\d+)\.(?:png|jpe?g)$', os.path.basename(path), re.I)
                        page_numbers.append(int(found.group(1)) if found else len(page_numbers) + 1)

            # 有文字层的 PDF 不会被 process_file 标记为扫描件，但封面、签章和日期
            # 仍必须通过页面图像核验。此时只渲染匹配文件的首页/末页，避免把整本报告
            # 的大量图片无差别发送给模型。
            if not page_imgs and file_info.get('extension') == '.pdf':
                is_target = not matched_file or file_info.get('filename', '') == matched_file
                if is_target:
                    try:
                        pdf = fitz.open(file_info.get('filepath', ''))
                        page_indexes = sorted(set([0, 1, len(pdf) - 2, len(pdf) - 1]))
                        page_imgs = []
                        page_numbers = []
                        for page_index in page_indexes:
                            if 0 <= page_index < len(pdf):
                                pix = pdf[page_index].get_pixmap(dpi=180)
                                page_imgs.append(pix.tobytes('jpg', jpg_quality=85))
                                page_numbers.append(page_index + 1)
                        pdf.close()
                    except Exception:
                        page_imgs = []
            if page_imgs:
                scanned_files.append({
                    'filename': file_info.get('filename', 'unknown'),
                    'page_images': page_imgs,
                    'page_numbers': page_numbers,
                })

        if not scanned_files:
            return self.review_single_checkpoint(checkpoint, file_texts, file_signatures, project_name, matched_file)

        prompt = self._build_single_prompt(checkpoint, file_texts, file_signatures, project_name, matched_file)

        content_items = [{
            "type": "text",
            "text": (
                f"{prompt}\n\n"
                f"【重要提醒】以下附上了相关文档的关键页面图片（JPEG 高质量格式，页码已标注）。"
                f"合同扫描件展示前17页和最后4页，其他长扫描件展示前12页和最后4页；文字型 PDF 展示首页、第二页、倒数第二页和末页。"
                f"请仔细查看已附页面中的文字、表格、签字盖章和日期；未附页面不能据此推断缺失。"
                f"对于封面/扉页/签章/编制人/审核人/批准人/日期检查，"
                f"必须以页面图片为准，不得因文本摘要未显示就认定缺失。"
            )
        }]

        # 按文件分组发送所有页面图片（无页数限制）
        for sf in scanned_files:
            if not sf['page_images']:
                continue
            content_items.append({
                "type": "text",
                "text": f"\n--- 文件: {sf['filename']}（本次展示原文件第 {', '.join(map(str, sf['page_numbers']))} 页）---"
            })
            for img_data in sf['page_images']:
                img_b64 = base64.b64encode(img_data).decode('utf-8')
                content_items.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}
                })

        try:
            response = self._create_completion(
                model=config.VOLC_VISION_MODEL,
                messages=[
                    {"role": "system", "content": self._get_system_prompt()},
                    {"role": "user", "content": content_items}
                ],
                temperature=0.1,
                max_tokens=2000,
            )
            result_text = self._extract_response_text(response)
            return self._parse_review_result(result_text, checkpoint)
        except Exception as e:
            return self.review_single_checkpoint(checkpoint, file_texts, file_signatures, project_name, matched_file)

    def _get_system_prompt(self):
        return """你是一个专业的电网数字化项目规范性审查专家。你的任务是：
1. 根据提供的检查要点，审查项目文件是否合规
2. 仔细阅读文件内容，对照检查要点逐项判断
3. 对于有"应提供""应包含""应有"等字眼的检查要点，必须确认文件中确实有相关内容，不能仅凭文件名判断
4. 对于日期类检查，需要核对时间先后顺序
5. 对于金额类检查，需要核对数值是否一致
6. **文件名容错**：实际文件名可能与检查要点中描述的文件名存在几个字的偏差（如"建设内容完成确认"与"项目建设内容完成确认单"），请根据文件实际内容判断是否对应，不要仅因文件名不完全匹配就判定"无法判断"或"不合格"
7. **图片范围**：合同扫描件展示前17页和最后4页，其他长扫描件展示前12页和最后4页，文字型PDF展示首页、第二页、倒数第二页和末页，正文内容应以提供的文本片段为准。未展示的页面不得被推断为缺失；需要未展示页面才能判断时，结论为“无法判断”。
8. **【重要】matched_file 必须填写**：你必须从下方的【项目资料文件清单】中选取**最相关的一个文件**作为 matched_file，返回其完整文件名（含扩展名如.pdf）。
    - 即使文件名不完全匹配，只要文件内容相关，就应当选取该文件
    - **禁止返回空字符串** - 每个检查要点通常都有对应的文件，请仔细判断并选取
    - 除非清单中确实**没有任何**文件与该检查要点相关，才返回空字符串
9. 当“对应文件类型”含有“【可能缺口材料】”或“【可能缺口材料，条件适用】”时，必须先从既有上传材料中寻找交叉证据。
   若现有材料不能明确佐证，result 必须为“无法判断”，不得仅因材料未上传判为“不合格”；
   但如果本检查项的【相关要求】明确规定该材料按本项目规则不上传且默认合格，必须遵守该例外，不得输出缺失问题；
   issue 必须说明无法交叉验证的具体事项，suggestion 才使用“人工审核建议：”开头，并明确写出要核验/补充的材料和字段。
10. 对“无法判断”的结论，不得猜测或补足金额、审批权限、系统留痕、签章、日期或评审过程。""

===== 签名盖章审查指引（请严格遵循以下推理步骤）=====

当检查要点涉及 **签字、盖章、签名、印章、签章、签到、审核、审批** 时，你必须按以下步骤逐一推理：

**步骤1：定位关键页面**
- 优先查看文档的 **最后两页**（尤其是最后一页），签名盖章通常出现在文件末尾的签章栏
- 如果最后两页没有找到，再往前翻阅前面的页面

**步骤2：查找签名位置**
- 在页面中查找以下关键位置字样旁边的签名：
  - 职务类：**"负责人"、"组长"、"经理"、"主任"、"院长"、"科长"、"主管"**
  - 动作类：**"签字"、"签章"、"签名"、"签署"、"审批"、"审核"、"审定"、"批准"**
  - 角色类：**"经办人"、"编制人"、"审核人"、"批准人"、"校核人"、"审查人"**
- **重点关注** 这些字样后面的下划线______或方框□位置，手写签名通常出现在这里

**步骤3：识别签字人姓名**
- 仔细查看手写签名区域，尝试辨认签名文字
- 如果签名潦草，可以结合签名下方的打印体姓名（如果有）来判断
- 如果实在无法辨认具体姓名，描述签名的样式（如：一个三字签名、潦草签字）

**步骤4：识别签字日期**
- 在签名位置的附近寻找日期信息（通常在签名的下一行或右侧）
- 日期格式通常为：2024年X月X日 或 2024-X-X 或 2024/X/X
- 注意区分：签署日期≠文件创建日期，以手写签名为准

**步骤5：识别盖章信息**
- 查看页面中是否有 **红色圆形公章** 或 **骑缝章**
- 尝试分辨公章上的文字（通常包含单位名称）
- 检查盖章是否覆盖了签名的部分

**步骤6：综合判断**
- 判断签名人数是否满足要求（如签到表需要多名专家签名）
- 判断签字/盖章位置是否完整
- 将以上发现的签名人、日期、印章等信息写入 evidence 字段

**如果签名无法清晰辨认，请在 evidence 中说明**：如"最后一页有手写签名位置，签名潦草无法辨认具体姓名"或"日期处有手写笔迹，但扫描件不够清晰"等。

输出格式要求（JSON）：
{
    "result": "合格" | "不合格" | "无法判断" | "不适用",
    "issue": "如果不合格，描述具体问题（如缺少谁签字、缺盖章、日期缺失等）；如果无法判断，说明缺少哪些信息",
    "suggestion": "整改建议",
    "evidence": "支撑判断的具体证据（引用文件内容；如果涉及签名盖章，必须写明发现了哪些签名、签字人、日期、印章信息）",
    "matched_file": "审查所依据的文件完整名称（必须从下方文件清单中选取，包含扩展名；非常重要，不要返回空字符串）"
}"""

    def _build_single_prompt(self, checkpoint, file_texts, file_signatures, project_name='', matched_file=''):
        """构建单条审查的提示词"""
        related_files_text = self._get_related_files(checkpoint, file_texts, file_signatures, matched_file)

        # 只展示本检查项候选文件，避免每条重复携带整个项目的文件清单。
        candidate_names = list(checkpoint.get('candidate_files', []))
        if matched_file and matched_file not in candidate_names:
            candidate_names.insert(0, matched_file)
        candidate_set = set(candidate_names)
        file_list_lines = []
        for fi in file_texts:
            file_key = fi.get('file_key', fi.get('filename', ''))
            if file_key in candidate_set:
                file_list_lines.append(f"  - {file_key}")
        file_list_str = '\n'.join(file_list_lines)

        # 如果预处理匹配已指定文件，标注出来
        matched_note = ''
        if matched_file:
            matched_note = f"\n【预处理匹配结果】本检查要点匹配的文件为: {matched_file}"

        prompt = f"""请审查以下项目文件：

【项目名称】{project_name}

【本检查项候选证据清单】（共 {len(file_list_lines)} 个文件；这是系统从完整项目档案中筛出的相关候选，不等同于全部上传文件）：
{file_list_str}
{matched_note}

【检查要点】
- 序号: {checkpoint['seq']}
- 项目阶段: {checkpoint['stage']}
- 项目环节: {checkpoint['link']}
- 对应文件类型: {checkpoint['doc_name']}
- 检查内容: {checkpoint['check_point']}
- 依据文件: {checkpoint.get('basis', '')}
- 相关要求: {checkpoint.get('requirement', '')}
- 历史风险点: {checkpoint.get('risk', '')}

【细则操作指导】
{checkpoint.get('requirement', '')}

【本项目取证定位与边界】
{checkpoint.get('stage_evidence_guide', '')}
{checkpoint.get('evidence_hint', '')}

【缺口材料判定规则】
        {('本检查项含可能缺口材料：先在既有上传材料中交叉取证；不能明确验证时，必须返回“无法判断”，并针对实际缺失内容生成一句具体建议；但如果【相关要求】明确规定该材料按本项目规则不上传且默认合格，必须遵守默认合格，不得输出缺失问题；不得机械复制细则中的建议原文，不得因未上传直接判定不合格。' if checkpoint.get('possible_gap_materials') else '本检查项未标注可能缺口材料，应依据已上传材料直接判断。')}

【项目资料文件内容】
{related_files_text}

请对照检查要点，审查上述文件是否合规，并严格按照以下要求输出：
1. matched_file 必须从上方【本检查项候选证据清单】中选取，返回完整文件名（含扩展名）
2. **即使文件名不完全匹配，只要文件内容相关，也要选取该文件**
3. **非常重要：不要返回空字符串！**
4. **证据边界**：候选清单是从完整项目档案中筛出的证据组，不是单一文件，也不是全部上传文件。必须阅读清单中的所有候选文件并做交叉核对，不能只看 matched_file；不能因为本消息未展示某文件就写“项目未上传”“未提供完整文件”或“全项目缺失”。只有在候选文件已实际阅读且能确认其缺页、缺附件或内容违反要求时，才能写“不合格”；否则写“在已检索候选资料中无法确认”，并给出针对性的人工核验动作。
5. **证据引用**：issue 中涉及事实时，必须写明至少一个“【证据：完整相对路径｜页码/章节】”；OCR文字可用于定位和交叉比对，印章、手写签名、颜色和复杂表格数值必须再以页面图片复核。不得把“未附视觉图片”表述为“文件未展示”。
6. suggestion 必须根据本次实际审查结果生成，简洁、明确、可执行：合格或不适用填写“无需整改”；不合格只写针对具体问题的整改动作；无法判断只写一条具体的“人工审核建议：……”，明确需要补充/核验的文件、页码、字段或时间关系。禁止只写“请补充材料”“建议人工审核”“人工审核建议”等空泛表述，建议控制在80字以内。

输出JSON格式结果。"""
        return prompt

    def _get_related_files(self, checkpoint, file_texts, file_signatures, matched_file=''):
        """根据检查要点获取相关文件内容
        如果预处理匹配已指定 matched_file，优先使用该文件
        """
        candidate_names = list(checkpoint.get('candidate_files', []))
        if matched_file and matched_file not in candidate_names:
            candidate_names.insert(0, matched_file)
        if candidate_names:
            related = []
            for name in candidate_names:
                for fi in file_texts:
                    if fi.get('file_key', fi.get('filename', '')) != name:
                        continue
                    text = self._retrieve_text(fi, checkpoint, max_chars=9000)
                    sig = fi.get('signature_stamp', {})
                    if fi.get('is_scanned'):
                        block = f"\n--- 候选文件: {name}（扫描件，请查看下方图片）---\n"
                        block += f"类型: {fi.get('file_type', '')}；完整文件页数: {fi.get('page_count', '未统计')}\n"
                        block += f"逐页OCR证据（以页面图片复核关键数字和签章）：\n{text}\n"
                    else:
                        block = f"\n--- 候选文件: {name} ---\n"
                        block += f"类型: {fi.get('file_type', '')}\n文本内容:\n{text}\n"
                    if sig.get('details'):
                        block += f"签名印章检测: {sig['details']}\n"
                    related.append(block)
                    break
            if related:
                return ''.join(related)

        doc_name = checkpoint.get('doc_name', '')
        link = checkpoint.get('link', '')
        stage = checkpoint.get('stage', '')

        related = []

        type_keywords = {
            '可研报告': ['可研报告', '可行性研究', '可研'],
            '可研评审意见': ['可研评审', '评审意见'],
            '可研批复文件': ['可研批复', '批复'],
            '合同': ['合同', '协议'],
            '中标': ['中标', '成交'],
            '采购': ['采购', '招标', '投标', '竞谈', '询价'],
            '需求规格': ['需求规格', '业务需求', '软件需求'],
            '概要设计': ['概要设计'],
            '详细设计': ['详细设计'],
            '测试': ['测试'],
            '实施方案': ['实施方案', '实施'],
            '验收': ['验收', '签到表', '专家组', '专家签到'],
            '结算': ['结算'],
            '决算': ['决算'],
            '转资': ['转资', '记账凭证'],
            '上线': ['上线', '试运行'],
            '用户报告': ['用户报告'],
            '竣工': ['竣工', '工作报告', '技术报告'],
            '安全防护': ['安全防护', '安全'],
            '应急预案': ['应急预案', '应急'],
            '用户手册': ['用户手册'],
            '管理员手册': ['管理员手册'],
            '部署方案': ['部署', '部署方案'],
            '数据字典': ['数据字典'],
        }

        relevant_types = []
        for kw, ftype in type_keywords.items():
            if kw in doc_name or kw in checkpoint.get('check_point', ''):
                relevant_types.append(ftype)

        for file_info in file_texts:
            filename = file_info.get('filename', '')
            file_type = file_info.get('file_type', '')

            matched = False
            for rtype in relevant_types:
                for kw in rtype:
                    if kw in filename or kw in file_type:
                        matched = True
                        break
                if matched:
                    break

            if matched or not relevant_types:
                text = file_info.get('text', '')
                sig = file_info.get('signature_stamp', {})
                page_imgs = file_info.get('page_images', [])

                if page_imgs:
                    file_block = f"\n--- 文件: {filename} (扫描件，请在下方图片中查看实际内容) ---\n"
                    file_block += f"类型: {file_type}\n"
                    file_block += f"此文件已作为图片附加在审查消息中，请直接查看图片识别文字、签名、印章和评分信息\n"
                else:
                    is_key_file = any(kw in filename for kw in ['签到', '签章', '验收意见', '签字', '盖章'])
                    max_len = 5000 if is_key_file else 3000
                    text = self._retrieve_text(file_info, checkpoint, max_chars=max_len)

                    file_block = f"\n--- 文件: {filename} ---\n"
                    file_block += f"类型: {file_type}\n"
                    file_block += f"文本内容:\n{text}\n"

                    if sig.get('details'):
                        file_block += f"签名印章检测: {sig['details']}\n"

                related.append(file_block)

        if not related:
            for file_info in file_texts:
                filename = file_info.get('filename', '')
                text = file_info.get('text', '')
                if len(text) > 500:
                    text = text[:500] + "..."
                related.append(f"\n--- 文件: {filename} ---\n{text}")

        return '\n'.join(related[:10])

    def _retrieve_text(self, file_info, checkpoint, max_chars=12000):
        """从页级证据库召回可定位片段，而不是仅截取文件开头。"""
        text = file_info.get('text', '') or ''
        knowledge_path = file_info.get('knowledge_text_path', '')
        if knowledge_path and os.path.exists(knowledge_path):
            try:
                with open(knowledge_path, 'r', encoding='utf-8') as f:
                    text = f.read()
            except OSError:
                pass
        query = f"{checkpoint.get('check_point', '')} {checkpoint.get('requirement', '')}"
        # 使用与文件索引一致的领域短语切分。原先的正则会将一整段连续中文
        # 操作指导当作一个词，全文检索几乎命不中，模型只能看到文首和文末片段。
        terms = _terms(query)

        units = []
        evidence_path = file_info.get('evidence_index_path', '')
        if evidence_path and os.path.exists(evidence_path):
            try:
                with open(evidence_path, 'r', encoding='utf-8') as f:
                    units = json.load(f)
            except (OSError, json.JSONDecodeError):
                units = []
        if not units:
            units = [{'file_key': file_info.get('file_key', file_info.get('filename', '')),
                      'locator': '全文', 'text': text}]

        scored = []
        for position, unit in enumerate(units):
            unit_text = unit.get('text', '')
            score = sum(unit_text.count(term) * max(2, len(term)) for term in terms[:25])
            # 始终保留首段，用于文件身份、目录、适用范围等上下文。
            if position == 0:
                score += 1
            if score:
                scored.append((score, position, unit))
        if not scored:
            scored = [(1, position, unit) for position, unit in enumerate(units[:3])]
        scored.sort(key=lambda item: (-item[0], item[1]))

        result = []
        used = 0
        for _, _, unit in scored[:12]:
            unit_text = unit.get('text', '').strip()
            if not unit_text:
                continue
            evidence = f"【证据定位：{unit.get('file_key', '')}｜{unit.get('locator', '')}】\n{unit_text}"
            if used + len(evidence) > max_chars:
                evidence = evidence[:max_chars - used]
            if evidence:
                result.append(evidence)
                used += len(evidence)
            if used >= max_chars:
                break
        return '\n\n[知识库相关片段]\n'.join(result)

    def _parse_review_result(self, result_text, checkpoint):
        """解析大模型返回的审查结果"""
        parsed = {}
        try:
            json_start = result_text.find('{')
            if json_start >= 0:
                # raw_decode 可在 JSON 后仍带解释文字时正确取出第一个对象，
                # 避免 rfind 把额外大括号一并纳入而解析失败。
                parsed, _ = json.JSONDecoder().raw_decode(result_text[json_start:])
        except (json.JSONDecodeError, TypeError, ValueError):
            parsed = {}

        if not isinstance(parsed, dict):
            parsed = {}
        issue = parsed.get('issue', '')
        if not issue:
            # 模型偶发输出格式接近JSON但不完整时，提取关键字段，避免把原始JSON展示给用户。
            match = re.search(r'"issue"\s*:\s*"((?:\\.|[^"\\])*)"', result_text, re.S)
            if match:
                try:
                    issue = json.loads(f'"{match.group(1)}"')
                except json.JSONDecodeError:
                    issue = match.group(1)
            else:
                issue = result_text[:200]

        return {
            'seq': checkpoint['seq'],
            'check_point': checkpoint['check_point'],
            'doc_name': checkpoint['doc_name'],
            'stage': checkpoint['stage'],
            'link': checkpoint['link'],
            'result': parsed.get('result', '无法判断'),
            'issue': issue,
            'suggestion': parsed.get('suggestion', ''),
            'evidence': parsed.get('evidence', ''),
            'matched_file': parsed.get('matched_file', ''),
        }
