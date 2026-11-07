# -*- coding: utf-8 -*-
"""
Excel 输出模块 - 将审查结果填入模板
"""
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import config


def _classify_problem(item):
    """根据审查要点和问题描述自动分类问题类型"""
    check_point = item.get('check_point', '')
    issue = item.get('issue', '')
    suggestion = item.get('suggestion', '')
    stage = item.get('stage', '')
    link = item.get('link', '')
    text = f"{check_point} {issue} {suggestion} {stage} {link}"

    # 材料缺失
    if any(kw in text for kw in ['缺失', '缺少', '未提供', '未提交', '缺', '没有提供', '未附']):
        return '材料缺失'
    # 项目成果与可研、合同不一致
    if any(kw in text for kw in ['不一致', '不符', '不匹配', '与合同', '与可研', '矛盾']):
        return '项目成果与可研、合同不一致'
    # 时序性错误
    if any(kw in text for kw in ['时序', '顺序错', '先后顺序', '时间', '逾期', '滞后', '超期']):
        return '时序性错误'
    # 材料内容错误或不规范
    if any(kw in text for kw in ['错误', '不规范', '不正确', '不完整', '不清晰', '格式', '模板', '不规范']):
        return '材料内容错误或不规范'
    return '其他'


def _write_problem_desc(item, original_result):
    """生成问题描述文本（只输出问题部分）"""
    issue = item.get('issue', '')
    return issue.strip() if issue else '无问题描述'


def _write_data_rows(ws, filtered_results, project_name, start_row,
                     content_font, wrap_alignment, thin_border,
                     red_fill, yellow_fill, green_fill,
                     col_map):
    """
    通用数据写入函数
    col_map: {col_number: ('field_name' or callable), ...}
    """
    current_row = start_row
    for item in filtered_results:
        original_result = item.get('original_result', item['result'])
        for col, spec in col_map.items():
            if callable(spec):
                value = spec(item, original_result)
            elif spec == 'seq':
                value = item.get('seq', 0)
            elif spec == 'project_name':
                value = project_name
            elif spec == 'doc_name':
                value = item.get('doc_name', '')
            elif spec == 'matched_file':
                value = item.get('matched_file', '') or item.get('doc_name', '')
            elif spec == 'problem_category':
                value = _classify_problem(item)
            elif spec == 'problem_desc':
                value = _write_problem_desc(item, original_result)
            elif spec == 'result':
                value = item['result']
            elif spec == 'stage':
                value = item.get('stage', '')
            elif spec == 'link':
                value = item.get('link', '')
            elif spec == 'check_point':
                value = item.get('check_point', '')
            elif spec == 'suggestion':
                value = item.get('suggestion', '')
            elif spec == 'issue':
                value = item.get('issue', '')
            elif spec == 'empty':
                value = ''
            else:
                value = spec  # 直接使用提供的值

            cell = ws.cell(row=current_row, column=col, value=value)
            cell.font = content_font
            cell.alignment = wrap_alignment
            cell.border = thin_border

        # 根据结果对问题描述列标色
        original_result = item.get('original_result', item['result'])
        result_fill = green_fill
        if original_result == '不合格':
            result_fill = red_fill
        elif original_result == '无法判断':
            result_fill = yellow_fill
        # 找问题描述列（col_map 中值为 'problem_desc' 的列）
        for col, spec in col_map.items():
            if spec == 'problem_desc':
                ws.cell(row=current_row, column=col).fill = result_fill
                break

        current_row += 1

    return current_row


def fill_excel_template(results, project_name, output_path=None, selected_stages=None, exclude_confirmed_seqs=None):
    """
    将审查结果填入第二季度规范性检查问题清单模板.xlsx

    只写入"无法判断"和"不合格"的审查项，
    手工复核确认合格的条目不写入（由 exclude_confirmed_seqs 控制）。

    同时写入"问题清单"和"问题台账"两个 Sheet，
    并将"问题台账"设为默认显示的 Sheet（用户打开 Excel 可直接看到数据）。
    """
    if exclude_confirmed_seqs is None:
        exclude_confirmed_seqs = set()

    if output_path is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(config.OUTPUT_FOLDER, f'审查结果_{project_name}_{timestamp}.xlsx')

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 复制模板
    if not os.path.exists(config.EXCEL_TEMPLATE):
        raise FileNotFoundError(f"模板文件不存在: {config.EXCEL_TEMPLATE}")
    shutil.copy(config.EXCEL_TEMPLATE, output_path)

    # 加载工作簿
    wb = load_workbook(output_path)

    # 已证实不合格才进入“问题”台账；“无法判断”是待补证/人工核验事项，
    # 不能被误写成项目问题或按关键词伪分类为时序错误、材料缺失等。
    def should_include_problem(item):
        if item.get('seq', 0) in exclude_confirmed_seqs:
            return False
        return item['result'] == '不合格'

    filtered_results = [r for r in results if should_include_problem(r)]
    pending_results = [
        r for r in results
        if r.get('seq', 0) not in exclude_confirmed_seqs and r['result'] == '无法判断'
    ]

    # 定义样式
    content_font = Font(name='微软雅黑', size=9)
    wrap_alignment = Alignment(wrap_text=True, vertical='top')
    center_alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_font = Font(name='微软雅黑', size=10, bold=True)
    header_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
    green_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')

    # ========== 写入"问题清单" Sheet ==========
    if '问题清单' in wb.sheetnames:
        ws1 = wb['问题清单']
    else:
        ws1 = wb[wb.sheetnames[0]]

    # 问题清单结构：
    # Row 1: 标题行（合并单元格）
    # Row 2: 分类标题
    # Row 3: 列标题（项目名称、合同名称、合同编号、问题描述...）
    # Row 4: 检查标准
    # Row 5~16: 示例数据

    # 取消数据区域中的合并单元格（Row 20+）
    for merged_range in list(ws1.merged_cells.ranges):
        if merged_range.min_row >= 20:
            try:
                ws1.unmerge_cells(str(merged_range))
            except Exception:
                pass

    # 从第20行开始追加审查结果汇总表（保留模板原有数据不受影响）
    start_row = 20
    # 先清空第20行及之后的数据
    for row in ws1.iter_rows(min_row=start_row, max_row=ws1.max_row, max_col=30):
        for cell in row:
            cell.value = None

    # 写入审查结果标题
    title_cell = ws1.cell(row=start_row, column=1, value=f'【审查结果汇总】项目名称: {project_name}')
    title_cell.font = Font(name='微软雅黑', size=12, bold=True)
    ws1.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
    review_start = start_row + 1

    # 写入表头（只写问题相关列，不写审查结果/建议）
    headers = ['序号', '审查阶段', '审查环节', '检查要点', '问题描述', '匹配文件']
    for col_idx, header in enumerate(headers, 1):
        cell = ws1.cell(row=review_start, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border

    # 写入数据（从 review_start + 1 行开始）
    current_row_qd = review_start + 1
    for item in filtered_results:
        seq = item.get('seq', '')
        stage = item.get('stage', '')
        link = item.get('link', '')
        check_point = item.get('check_point', '')
        issue = item.get('issue', '')
        matched_file = item.get('matched_file', '') or item.get('doc_name', '')

        row_data = [seq, stage, link, check_point, issue, matched_file]
        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=current_row_qd, column=col_idx, value=val)
            cell.font = content_font
            cell.alignment = wrap_alignment if col_idx >= 4 else center_alignment
            cell.border = thin_border

        # 根据原结果标色
        original_result = item.get('original_result', item['result'])
        if original_result == '不合格':
            for col_idx in range(1, 7):
                ws1.cell(row=current_row_qd, column=col_idx).fill = red_fill
        elif original_result == '无法判断':
            for col_idx in range(1, 7):
                ws1.cell(row=current_row_qd, column=col_idx).fill = yellow_fill

        current_row_qd += 1

    # 如果没有数据，填入提示
    if current_row_qd == review_start + 1:
        cell = ws1.cell(row=current_row_qd, column=1, value='经审查，未发现明显规范性问题。')
        cell.font = content_font
        cell.alignment = wrap_alignment
        cell.border = thin_border
        current_row_qd += 1

    # 设置列宽（审查结果区域）
    col_widths_qd = {1: 8, 2: 14, 3: 14, 4: 50, 5: 50, 6: 25}
    for col, width in col_widths_qd.items():
        try:
            ws1.column_dimensions[get_column_letter(col)].width = width
        except Exception:
            pass

    # ========== 写入"问题台账" Sheet ==========
    if '问题台账' in wb.sheetnames:
        ws2 = wb['问题台账']
    else:
        ws2 = wb[wb.sheetnames[0]]

    # 问题台账结构：
    # Row 1: 列标题（序号、项目编码、项目名称、合同名称、项目部门、责任人、问题分类、问题材料名称、问题描述...）
    # Row 2~101: 数据行（已有序号 1~100）

    # 清空第2行开始的全部数据（保留第1行表头）
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=ws2.max_column):
        for cell in row:
            cell.value = None

    # 写入数据到问题台账（从第2行开始）
    # 列映射：A=序号, C=项目名称, D=合同名称, G=问题分类, H=问题材料名称, I=问题描述
    taizhang_col_map = {
        1: 'seq',                 # A: 序号
        3: 'project_name',        # C: 项目名称
        4: 'doc_name',            # D: 合同名称
        7: 'problem_category',    # G: 问题分类
        8: 'matched_file',        # H: 问题材料名称
        9: 'problem_desc',        # I: 问题描述
    }
    current_row_tz = _write_data_rows(
        ws2, filtered_results, project_name, 2,
        content_font, wrap_alignment, thin_border,
        red_fill, yellow_fill, green_fill,
        taizhang_col_map
    )

    # 如果没有数据，填入提示
    if current_row_tz == 2:
        ws2.cell(row=2, column=3, value=project_name)
        cell = ws2.cell(row=2, column=9, value='经审查，未发现明显规范性问题。')
        cell.font = content_font
        cell.alignment = wrap_alignment
        cell.border = thin_border
        current_row_tz = 3

    # 调整列宽（问题台账前9列）
    col_widths_tz = {1: 6, 2: 15, 3: 25, 4: 25, 5: 12, 6: 12, 7: 20, 8: 30, 9: 60}
    for col, width in col_widths_tz.items():
        try:
            ws2.column_dimensions[get_column_letter(col)].width = width
        except Exception:
            pass

    # ========== 写入“待人工核验” Sheet ===========
    if '待人工核验' in wb.sheetnames:
        del wb['待人工核验']
    pending_ws = wb.create_sheet('待人工核验')
    pending_headers = ['序号', '审查阶段', '检查要点', '待核验原因', '建议核验动作', '已匹配文件']
    for col_idx, header in enumerate(pending_headers, 1):
        cell = pending_ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = thin_border
    for row_idx, item in enumerate(pending_results, 2):
        values = [
            item.get('seq', ''), item.get('stage', ''), item.get('check_point', ''),
            item.get('issue', ''), item.get('suggestion', ''),
            item.get('matched_file', '') or item.get('doc_name', ''),
        ]
        for col_idx, value in enumerate(values, 1):
            cell = pending_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = content_font
            cell.alignment = wrap_alignment if col_idx >= 3 else center_alignment
            cell.border = thin_border
            cell.fill = yellow_fill
        pending_ws.row_dimensions[row_idx].height = 60
    for col, width in enumerate((8, 14, 45, 60, 50, 45), 1):
        pending_ws.column_dimensions[get_column_letter(col)].width = width

    # 调整行高
    for row in range(2, current_row_tz):
        try:
            ws2.row_dimensions[row].height = 60
        except Exception:
            pass

    # 将"问题台账"设为激活（默认显示）的 Sheet
    if '问题台账' in wb.sheetnames:
        wb.active = wb.sheetnames.index('问题台账')

    # 保存
    wb.save(output_path)
    print(f"审查结果已保存到: {output_path}")
    print(f"  数据总量: {len(results)} 条, 已证实问题: {len(filtered_results)} 条, 待人工核验: {len(pending_results)} 条")
    print(f"  问题清单 (第{review_start+1}行起): {current_row_qd - review_start - 1} 条")
    print(f"  问题台账: {current_row_tz - 2} 条")
    return output_path


def generate_simplified_report(results, project_name, reviewer=None, output_path=None):
    """
    手动复核确认后，调用 LLM 生成简化版审查报告（XLSX）
    包含所有审查项（合格/不合格/无法判断/不适用），简明扼要
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    if output_path is None:
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(config.OUTPUT_FOLDER, f'简化审查报告_{project_name}_{timestamp}.xlsx')

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = '审查结果'

    # 列宽设置
    col_widths = {'A': 8, 'B': 15, 'C': 15, 'D': 50, 'E': 12, 'F': 40, 'G': 40}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # 样式定义
    title_font = Font(name='微软雅黑', size=14, bold=True)
    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    content_font = Font(name='微软雅黑', size=10)
    header_fill = PatternFill(start_color='1a5276', end_color='1a5276', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    content_alignment = Alignment(vertical='top', wrap_text=True)
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    pass_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
    fail_fill = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')
    unknown_fill = PatternFill(start_color='FCF3CF', end_color='FCF3CF', fill_type='solid')

    # 第1行：标题
    ws.merge_cells('A1:G1')
    title_cell = ws['A1']
    title_cell.value = f'项目规范性审查简化报告 - {project_name}'
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    # 第2行：统计
    total = len(results)
    qualified = sum(1 for r in results if r['result'] == '合格')
    unqualified = sum(1 for r in results if r['result'] == '不合格')
    unknown = sum(1 for r in results if r['result'] == '无法判断')
    na = sum(1 for r in results if r['result'] == '不适用')

    ws.merge_cells('A2:G2')
    stat_cell = ws['A2']
    stat_cell.value = f'总检查项: {total}  |  合格: {qualified}  |  不合格: {unqualified}  |  无法判断: {unknown}  |  不适用: {na}'
    stat_cell.font = Font(name='微软雅黑', size=10, bold=True)
    stat_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 25

    # 第3行：表头
    headers = ['序号', '项目阶段', '审查环节', '检查要点', '审查结果', '问题/说明', '整改建议']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[3].height = 30

    # 尝试调用 LLM 生成简化描述（如果 reviewer 可用）
    simplified_descriptions = {}
    if reviewer and total > 0:
        try:
            # 分批调用 LLM 生成简化摘要
            batch_size = 10
            for i in range(0, len(results), batch_size):
                batch = results[i:i+batch_size]
                batch_text = '\n'.join([
                    f"[{r['seq']}] {r['check_point'][:60]} -> {r['result']}"
                    for r in batch
                ])
                llm_prompt = f"""你是一个项目审查报告撰写专家。请为以下项目审查结果生成简化的单行描述（每个条目一行，简洁明了，直接指出关键问题或说明已通过）。

项目名称: {project_name}

审查结果:
{batch_text}

请按以下格式输出（每条一行，不要序号）：
[条目序号] 简短描述（20-30字，说明关键问题或"符合要求"）

例如：
[1] 可研报告内容完整，符合要求
[2] 合同缺少双方法定代表人签字"""
                try:
                    response = reviewer.client.chat.completions.create(
                        model=config.VOLC_VISION_MODEL,
                        messages=[
                            {"role": "system", "content": "你是一个专业、简洁的报告撰写助手，请用中文输出。"},
                            {"role": "user", "content": [{"type": "text", "text": llm_prompt}]}
                        ],
                        temperature=0.3,
                        max_tokens=1000,
                    )
                    resp_text = reviewer._extract_response_text(response)
                    for line in resp_text.strip().split('\n'):
                        line = line.strip()
                        if line.startswith('[') and ']' in line:
                            try:
                                seq_str = line[1:line.index(']')]
                                seq = int(seq_str)
                                desc = line[line.index(']')+1:].strip()
                                simplified_descriptions[seq] = desc
                            except (ValueError, IndexError):
                                pass
                except Exception:
                    pass
        except Exception:
            pass

    # 填写数据
    row_fills = {
        '合格': pass_fill,
        '不合格': fail_fill,
        '无法判断': unknown_fill,
        '不适用': PatternFill(start_color='EBF5FB', end_color='EBF5FB', fill_type='solid'),
    }

    for idx, r in enumerate(results):
        row = idx + 4
        ws.row_dimensions[row].height = 50

        seq = r['seq']
        result = r['result']

        values = [
            seq,
            r.get('stage', ''),
            r.get('link', ''),
            r.get('check_point', ''),
            result,
            simplified_descriptions.get(seq, r.get('issue', '')),
            simplified_descriptions.get(seq, r.get('suggestion', '')),
        ]

        fill = row_fills.get(result, PatternFill())

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = content_font
            cell.alignment = content_alignment if col_idx > 3 else center_alignment
            cell.border = thin_border
            if col_idx == 5:  # 结果列用颜色填充
                cell.fill = fill
            cell.fill = fill

    wb.save(output_path)
    print(f"简化报告已保存到: {output_path}")
    return output_path
