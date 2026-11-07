# -*- coding: utf-8 -*-
"""
主 Agent 编排器 - 串联所有模块
"""
import os
import shutil
import hashlib
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

from pdf_processor import process_file, convert_key_pages_to_images, _guess_file_type
from reviewer import Reviewer
from excel_writer import fill_excel_template
from preliminary_rules import load_rule_checkpoints
from knowledge_base import create_knowledge_dir, save_document_text, save_evidence_index, cleanup_knowledge_dir
from file_index import select_candidates
from archive_extractor import extract_supported_documents
import config


def _write_review_log(lines, path, append=False):
    """把流程日志打印到控制台并写入 TXT；控制台编码异常不影响审查。"""
    with open(path, 'a' if append else 'w', encoding='utf-8') as log_file:
        for line in lines:
            log_file.write(line + '\n')
            try:
                print(line)
            except UnicodeEncodeError:
                print(line.encode('ascii', errors='replace').decode('ascii'))


class ReviewAgent:
    """规范性审查 Agent"""

    @staticmethod
    def available_stages():
        """返回附件2已生成检查要点的阶段，按常用流程顺序展示。"""
        preferred_order = ['项目前期', '项目采购', '项目建设', '上下线', '项目验收', '决算', '转资']
        if not os.path.isdir(config.CHECKPOINTS_FOLDER):
            return []
        existing = {
            os.path.splitext(name)[0]
            for name in os.listdir(config.CHECKPOINTS_FOLDER)
            if name.endswith('.txt')
        }
        # 目录中可能存在修改说明、人工备注等 TXT，不能把它们当作阶段规则加载。
        return [stage for stage in preferred_order if stage in existing]

    def __init__(self, api_key=None):
        self.reviewer = Reviewer(api_key=api_key)
        self.checkpoints = None
        self.selected_stages = None
        self.project_name = ''
        self.project_type = '开发实施项目'
        self.knowledge_dir = None

    def use_stage_rules(self, stages=None):
        """按选择加载附件2生成的检查要点，可同时加载多个阶段。"""
        allowed = {stage: os.path.join(config.CHECKPOINTS_FOLDER, f'{stage}.txt') for stage in self.available_stages()}
        stages = [s for s in (stages or self.available_stages()) if s in allowed]
        if not stages:
            stages = list(allowed)
        items = []
        for stage in stages:
            items.extend(load_rule_checkpoints(allowed[stage], stage))
        for seq, item in enumerate(items, 1):
            item['seq'] = seq
        self.checkpoints = {'开发实施项目': items}
        self.project_type = '开发实施项目'
        return self.checkpoints

    def set_stages(self, stages):
        """设置需要审查的环节"""
        self.selected_stages = stages

    def process_zip(self, zip_path, task_id=None, selected_filenames=None, metadata_only=False, extract_dir=None):
        """解压并处理 ZIP；metadata_only 只建立文件清单，不读取文件内容。"""
        if not os.path.exists(zip_path):
            raise FileNotFoundError(f"ZIP文件不存在: {zip_path}")

        # 只解压一次；第二阶段复用同一个目录，避免大 ZIP 重复解压。
        if extract_dir is None:
            extract_dir = os.path.join(config.EXTRACT_FOLDER,
                                       datetime.now().strftime('%Y%m%d_%H%M%S'))
            os.makedirs(extract_dir, exist_ok=True)
            extract_supported_documents(zip_path, extract_dir)
        elif not os.path.isdir(extract_dir) or not any(os.scandir(extract_dir)):
            # 调用方指定新目录时也必须执行解压；否则 metadata_only 会得到空文件
            # 清单，继而把所有检查项错误写成“未检索到”。已有解压目录则继续复用。
            os.makedirs(extract_dir, exist_ok=True)
            extract_supported_documents(zip_path, extract_dir)

        # 获取项目名称
        self.project_name = os.path.splitext(os.path.basename(zip_path))[0]

        # 图片存储目录（用于后续手动复核）
        images_dir = None
        if task_id:
            images_dir = os.path.join(config.EXTRACT_FOLDER, f'images_{task_id}')
            os.makedirs(images_dir, exist_ok=True)
            self.knowledge_dir = create_knowledge_dir(config.KNOWLEDGE_FOLDER, task_id)

        # 遍历所有文件
        file_infos = []
        for root, dirs, files in os.walk(extract_dir):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in config.SUPPORTED_EXTENSIONS:
                    file_path = os.path.join(root, f)
                    file_key = os.path.relpath(file_path, extract_dir).replace('\\', '/')
                    if selected_filenames and file_key not in selected_filenames:
                        continue
                    if metadata_only:
                        file_infos.append({
                            'filename': f,
                            'filepath': file_path,
                            'relative_path': file_key,
                            'file_key': file_key,
                            'extension': ext,
                            'text': '',
                            'page_images': [],
                            'page_image_paths': [],
                            'signature_stamp': {},
                            'file_type': _guess_file_type(f),
                        })
                        continue
                    try:
                        info = process_file(file_path)
                        info['relative_path'] = file_key
                        info['file_key'] = file_key
                        # 任务结束后会清理解压目录；为手动复核保留匹配到的 PDF 原文件，
                        # 之后按页懒渲染，既能查看全文，也不会在审查阶段把整本 PDF 全部转图。
                        if task_id and images_dir and ext == '.pdf':
                            source_dir = os.path.join(images_dir, 'sources')
                            os.makedirs(source_dir, exist_ok=True)
                            source_id = hashlib.sha1(file_key.encode('utf-8')).hexdigest()[:16]
                            preview_source = os.path.join(source_dir, f'{source_id}.pdf')
                            if not os.path.exists(preview_source):
                                try:
                                    os.link(file_path, preview_source)
                                except OSError:
                                    shutil.copy2(file_path, preview_source)
                            info['preview_source_path'] = preview_source
                        # 长文本写入任务专属临时知识库，Reviewer 按检查项检索，不把全文一次性塞进提示词。
                        if self.knowledge_dir and info.get('text'):
                            info['knowledge_text_path'] = save_document_text(
                                self.knowledge_dir, f, info['text']
                            )
                            info['evidence_index_path'] = save_evidence_index(
                                self.knowledge_dir, file_key, info['text']
                            )
                        # 如果有task_id，将页面图片保存到磁盘并用路径替换bytes
                        if task_id and info.get('page_images'):
                            saved_paths = []
                            for idx, img_data in enumerate(info['page_images']):
                                img_filename = f"{os.path.splitext(f)[0]}_page_{idx+1}.png"
                                img_path = os.path.join(images_dir, img_filename)
                                with open(img_path, 'wb') as imgf:
                                    imgf.write(img_data)
                                saved_paths.append(img_path)
                            info['page_image_paths'] = saved_paths
                            # 保留第一张图片用于预览（small amount of data）
                            # 删除bytes数据，用路径替代
                            info['page_images'] = []  # 清空bytes数据
                        elif task_id and info.get('extension') == '.pdf':
                            # 文字型PDF也保留关键页面图片，供封面、签章和日期检查使用。
                            saved_paths = []
                            for page_index, img_data in convert_key_pages_to_images(file_path):
                                img_filename = f"{os.path.splitext(f)[0]}_page_{page_index + 1}.jpg"
                                img_path = os.path.join(images_dir, img_filename)
                                with open(img_path, 'wb') as imgf:
                                    imgf.write(img_data)
                                saved_paths.append(img_path)
                            info['page_image_paths'] = saved_paths
                        file_infos.append(info)
                    except Exception as e:
                        print(f"处理文件失败 {f}: {e}")

        self.file_infos = file_infos
        return file_infos, extract_dir

    def run_review(self, zip_path, selected_stages=None, progress_callback=None, task_id=None):
        """
        执行完整审查流程
        返回: (results, excel_path, report_path, file_infos)
        """
        if progress_callback:
            progress_callback('正在加载检查要点...', 5)

        # 1. 加载检查要点
        if not self.checkpoints:
            self.use_stage_rules()

        if progress_callback:
            progress_callback('正在解压和处理文件...', 10)

        # 流程日志：记录第一遍解压清单和每条检查要点的本地候选文件。
        os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)
        log_path = os.path.join(
            config.OUTPUT_FOLDER,
            f'审查流程日志_{task_id or datetime.now().strftime("%Y%m%d_%H%M%S")}.txt',
        )

        # 2. 先解压并只建立文件名/路径清单，不读取无关文件内容。
        metadata_infos, extract_dir = self.process_zip(zip_path, metadata_only=True)

        log_lines = [
            '=' * 80,
            f'第一遍解压文件清单（共 {len(metadata_infos)} 个受支持文件）',
            f'解压目录：{extract_dir}',
            f'日志文件：{log_path}',
            '-' * 80,
        ]
        for index, info in enumerate(metadata_infos, 1):
            try:
                size = os.path.getsize(info.get('filepath', ''))
            except OSError:
                size = 0
            log_lines.append(
                f'{index:>4}. [{info.get("file_type", "")}] {info.get("relative_path", "")}  ({size} 字节)'
            )
        if progress_callback:
            progress_callback(f'已建立 {len(metadata_infos)} 个文件索引，准备匹配', 20)

        # 3. 筛选检查项
        if selected_stages:
            self.set_stages(selected_stages)

        filtered_checkpoints = self._filter_checkpoints()

        # 把解析后的检查项完整字段并入首段日志，与文件清单一次性写入。
        log_lines.extend([
            '',
            '=' * 80,
            f'解析后的结构化检查项（共 {len(filtered_checkpoints)} 条）',
            '-' * 80,
        ])
        for checkpoint in filtered_checkpoints:
            log_lines.append(
                f'[{checkpoint.get("seq")}] {json.dumps(checkpoint, ensure_ascii=False, default=str)}'
            )
        _write_review_log(log_lines, log_path)

        if progress_callback:
            progress_callback(f'筛选出 {len(filtered_checkpoints)} 条审查要点', 25)

        # 4. 预处理匹配：为每个检查要点匹配最相关的文件
        if progress_callback:
            progress_callback('正在执行预处理文件匹配...', 28)

        # 项目档案目录统一，优先根据附件2中的建议检索路径和文件名完成本地匹配。
        # 不再为第一轮匹配发送模型请求，也不读取无关文件的首页，避免大资料包长时间停在“匹配”。
        initial_candidates = {
            checkpoint['seq']: select_candidates(
                checkpoint, metadata_infos, limit=config.MAX_CANDIDATES_PER_CHECKPOINT
            )
            for checkpoint in filtered_checkpoints
        }
        matched_files = {
            checkpoint['seq']: initial_candidates.get(checkpoint['seq'], [''])[0]
            if initial_candidates.get(checkpoint['seq']) else ''
            for checkpoint in filtered_checkpoints
        }
        matched_count = sum(1 for v in matched_files.values() if v)
        if progress_callback:
            progress_callback(f'文件匹配完成: {matched_count}/{len(filtered_checkpoints)} 条已匹配', 30)

        candidate_map = {}
        candidate_details_map = {}
        for checkpoint in filtered_checkpoints:
            candidates, details = select_candidates(
                checkpoint, metadata_infos, matched_files.get(checkpoint['seq'], ''),
                limit=config.MAX_CANDIDATES_PER_CHECKPOINT,
                with_details=True,
            )
            candidate_map[checkpoint['seq']] = candidates
            candidate_details_map[checkpoint['seq']] = details
            if matched_files.get(checkpoint['seq'], '') not in candidates:
                matched_files[checkpoint['seq']] = candidates[0] if candidates else ''

        log_lines = [
            '',
            '=' * 80,
            f'每条检查要点的本地候选文件（共 {len(filtered_checkpoints)} 条）',
            '-' * 80,
        ]
        for checkpoint in filtered_checkpoints:
            seq = checkpoint['seq']
            candidates = candidate_map.get(seq, [])
            log_lines.append(
                f'[{seq}] {checkpoint.get("stage", "")}/{checkpoint.get("link", "")} | '
                f'{checkpoint.get("check_point", "")}'
            )
            log_lines.append(f'     主匹配：{matched_files.get(seq) or "（无）"}')
            if not candidates:
                log_lines.append('     候选：无')
            details = candidate_details_map.get(seq, {})
            for index, name in enumerate(candidates, 1):
                detail = details.get(name, {})
                log_lines.append(
                    f'     {index}. [{detail.get("score", 0)}分 | {detail.get("source", "")}] {name}'
                )
        _write_review_log(log_lines, log_path, append=True)

        selected_filenames = {name for names in candidate_map.values() for name in names}
        file_infos, extract_dir = self.process_zip(
            zip_path, task_id=task_id, selected_filenames=selected_filenames,
            extract_dir=extract_dir
        )
        if progress_callback:
            progress_callback(f'已读取匹配文件 {len(file_infos)} 个，无关文件已跳过', 32)

        # 5. 逐条审查
        results = []
        total = len(filtered_checkpoints)
        task_extract_dir = extract_dir

        def review_one(index, raw_checkpoint):
            matched_file = matched_files.get(raw_checkpoint['seq'], '')
            cp = {**raw_checkpoint, 'candidate_files': candidate_map.get(raw_checkpoint['seq'], [])}

            # 判断是否需要 Vision 审查
            cp_text = cp.get('check_point', '')
            doc_name = cp.get('doc_name', '')
            keyword_vision = any(kw in cp_text for kw in [
                '签字', '盖章', '签名', '印章', '公章', '签章', '签到',
                '封面', '扉页', '编制人', '审核人', '批准人', '最终版', '日期'
            ])
            keyword_vision = keyword_vision or any(kw in doc_name for kw in ['签到', '签章', '验收意见'])

            # 只有匹配文件本身是扫描件时才启用视觉；数字PDF的关键页图片只用于
            # 封面/签章/日期等明确需要视觉核验的检查，避免所有检查项重复传图。
            # 金额、时序、主体等检查经常同时依赖“文字型报价表/批复”和“扫描版合同/通知书”。
            # 不能只看模型选中的单一主文件；任一候选为扫描件时都要启用视觉审查，
            # 否则模型会把已上传的合同、中标通知书误判成“未提供”。
            candidate_keys = set(cp.get('candidate_files', []))
            has_scanned = any(
                item.get('file_key', item.get('filename', '')) in candidate_keys
                and item.get('is_scanned')
                for item in file_infos
            )

            # 有扫描件(PDF含图片) 或 关键词命中时，使用视觉审查（发送图片）
            # DOC/DOCX 文件无图片，自动使用纯文本审查
            if has_scanned or keyword_vision:
                result = self.reviewer.review_with_vision(
                    cp, file_infos,
                    {f['filename']: f.get('signature_stamp', {}) for f in file_infos},
                    self.project_name,
                    vision_model='doubao',
                    matched_file=matched_file
                )
            else:
                result = self.reviewer.review_single_checkpoint(
                    cp, file_infos,
                    {f['filename']: f.get('signature_stamp', {}) for f in file_infos},
                    self.project_name,
                    matched_file=matched_file
                )
            # matched_file 仅用于标识最主要的材料；保留完整候选证据组，供结果展示、
            # 导出和人工复核使用，避免一条要点被错误地收缩为一个文件。
            result['candidate_files'] = list(cp['candidate_files'])
            return index, result

        # 两路并发可消除网络等待时间，又避免大量视觉请求同时挤占模型配额。
        completed = 0
        with ThreadPoolExecutor(max_workers=config.REVIEW_CONCURRENCY) as executor:
            futures = [
                executor.submit(review_one, index, checkpoint)
                for index, checkpoint in enumerate(filtered_checkpoints)
            ]
            ordered_results = [None] * total
            for future in as_completed(futures):
                index, result = future.result()
                ordered_results[index] = result
                completed += 1
                if progress_callback:
                    progress_pct = 32 + int(completed / total * 53)
                    progress_callback(
                        f'正在审查 [{completed}/{total}]：{result["check_point"][:40]}...',
                        progress_pct,
                    )
        results = ordered_results

        if progress_callback:
            progress_callback('正在生成报告...', 85)

        # 5. 生成 Excel；临时文字知识库在本次任务结束后删除。
        os.makedirs(config.OUTPUT_FOLDER, exist_ok=True)
        try:
            excel_path = fill_excel_template(results, self.project_name,
                                             selected_stages=selected_stages)
            report_path = ''
            return results, excel_path, report_path, file_infos
        finally:
            cleanup_knowledge_dir(self.knowledge_dir)
            self.knowledge_dir = None
            if task_extract_dir and os.path.isdir(task_extract_dir):
                shutil.rmtree(task_extract_dir, ignore_errors=True)
            if progress_callback:
                progress_callback('审查完成!', 100)

    def _filter_checkpoints(self):
        """根据选择的环节和项目类型筛选检查要点"""
        if not self.checkpoints:
            return []

        if self.project_type in self.checkpoints:
            items = self.checkpoints[self.project_type]
        else:
            items = self.checkpoints.get('开发实施项目', [])

        if not self.selected_stages:
            return items

        filtered = []
        for item in items:
            if item['link'] in self.selected_stages or item['stage'] in self.selected_stages:
                filtered.append(item)
        return filtered

