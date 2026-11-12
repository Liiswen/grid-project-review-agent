# -*- coding: utf-8 -*-
"""材料完整性检查模块：根据检查文件清单.xlsx 判断各阶段材料是否齐全。"""

import os
import re
import openpyxl
import config


_HEADER_FIELDS = {
    '项目阶段': 'stage',
    '文档名称': 'doc_name',
    '别名': 'alias',
    '存放路径': 'storage_path',
    '是否有路径即判定为存在': 'path_exists_only',
}

_TRUE_FLAGS = {'是', '有', 'y', 'yes', 'true', '1', '√', '✓'}


def _normalize_header(value):
    """规范化表头文本：去除空格、全角空格。"""
    if value is None:
        return ''
    return str(value).replace(' ', '').replace('　', '').strip()


def _is_yes(value):
    """判断“是否有路径即判定为存在”字段是否为是。"""
    if value is None:
        return False
    v = str(value).replace(' ', '').replace('　', '').strip().lower()
    return v in _TRUE_FLAGS


def split_aliases(alias):
    """拆分别名字段，支持顿号、逗号、分号、斜杠等分隔符。

    括号内的分隔符不参与拆分，避免复合名称被错误切开。
    """
    if not alias:
        return []
    text = str(alias)
    parts = []
    buf = ''
    depth = 0
    for ch in text:
        if ch in '（(':
            depth += 1
        elif ch in '）)':
            depth = max(0, depth - 1)
        if depth == 0 and ch in '、，,;；/｜|\\\n\r\t':
            if buf.strip():
                parts.append(buf.strip())
            buf = ''
            continue
        buf += ch
    if buf.strip():
        parts.append(buf.strip())
    return parts


def load_material_manifest(excel_path=None):
    """读取检查文件清单.xlsx，返回按阶段分组的文档清单。

    支持五列结构：项目阶段 / 文档名称 / 别名 / 存放路径 / 是否有路径即判定为存在；
    缺少别名、判定列时（三列结构）按缺失处理。列位置优先按表头名称识别，
    表头无法识别时回退为按位置解析。

    返回结构：
    [
        {
            'stage': '项目前期',
            'doc_name': '业务需求报告',
            'alias': '需求报告、业务需求',
            'aliases': ['需求报告', '业务需求'],
            'storage_path': '00 项目需求',
            'path_exists_only': False,
        },
        ...
    ]
    阶段为空的行继承上一行的阶段。
    """
    if excel_path is None:
        excel_path = os.path.join(config.BASE_DIR, '检查文件清单.xlsx')
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f'检查文件清单不存在: {excel_path}')

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close()
        return []

    # 识别表头列位置，识别不到时回退为位置解析
    header = [_normalize_header(cell) for cell in rows[0]]
    col_index = {}
    for idx, name in enumerate(header):
        field = _HEADER_FIELDS.get(name)
        if field and field not in col_index:
            col_index[field] = idx
    if 'stage' not in col_index:
        col_index = {'stage': 0, 'doc_name': 1, 'storage_path': 2}

    def _cell(row, field):
        idx = col_index.get(field)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    # 跳过表头
    manifest = []
    current_stage = None
    for row in rows[1:]:
        if not row or all(cell is None for cell in row):
            continue
        stage = _cell(row, 'stage')
        doc_name = _cell(row, 'doc_name')
        alias = _cell(row, 'alias')
        storage_path = _cell(row, 'storage_path')
        path_flag = _cell(row, 'path_exists_only')

        if stage:
            current_stage = str(stage).strip()
        if not doc_name:
            continue
        alias_text = str(alias).strip() if alias else ''
        manifest.append({
            'stage': current_stage,
            'doc_name': str(doc_name).strip(),
            'alias': alias_text,
            'aliases': split_aliases(alias_text),
            'storage_path': str(storage_path).strip() if storage_path else None,
            'path_exists_only': _is_yes(path_flag),
        })
    wb.close()
    return manifest


def _strip_brackets(text):
    """去除中英文括号及其内部内容。"""
    # 去除中文括号及内容
    text = re.sub(r'（[^）]*）', '', text)
    # 去除英文括号及内容
    text = re.sub(r'\([^)]*\)', '', text)
    return text


def _normalize_path_for_match(storage_path):
    """规范化存放路径用于匹配。

    规则：去除括号及括号内内容 → 去除所有空格 → 统一 / 分隔符。
    """
    if not storage_path:
        return ''
    value = _strip_brackets(storage_path)
    value = value.replace(' ', '').replace('　', '')
    value = value.replace('\\', '/').rstrip('/')
    return value


def _path_matches(relative_path, normalized_storage_path):
    """判断文件相对路径是否命中规范化后的存放路径。

    文件路径同样去除括号及空格后，判断存放路径是否为其子串。
    """
    if not normalized_storage_path:
        return False
    rel = _strip_brackets(relative_path)
    rel = rel.replace(' ', '').replace('　', '').replace('\\', '/')
    return normalized_storage_path in rel


def _normalize_name(text):
    """规范化名称用于匹配：去扩展名、去空格。"""
    if not text:
        return ''
    value = os.path.splitext(str(text))[0]
    return value.replace(' ', '').replace('　', '')


def _name_matches(filename, doc_name):
    """判断文件名是否包含文档名称（双向包含，提高容错）。"""
    if not filename or not doc_name:
        return False
    fname = _normalize_name(filename)
    name = _normalize_name(doc_name)
    return name in fname or fname in name


def _alias_matches(filename, aliases):
    """判断是否有别名完全落入文件名中。"""
    if not filename or not aliases:
        return False
    fname = _normalize_name(filename)
    for alias in aliases:
        name = _normalize_name(alias)
        if name and name in fname:
            return True
    return False


def _matches_doc(filename, doc_name, aliases):
    """文件名命中判断：文档名称或任一别名命中即为命中。"""
    return _name_matches(filename, doc_name) or _alias_matches(filename, aliases)


def check_material_completeness(metadata_infos, manifest=None, excel_path=None):
    """根据清单检查材料完整性。

    Args:
        metadata_infos: process_zip(metadata_only=True) 返回的文件清单
        manifest: 文档清单，若为 None 则从 excel_path 读取
        excel_path: 检查文件清单路径

    Returns:
        (material_map, completeness_report)
        - material_map: {(stage, doc_name): {'exists': bool, 'file_key': str, 'filepath': str}}
        - completeness_report: 按阶段分组的完整性报告，用于前端展示
    """
    if manifest is None:
        manifest = load_material_manifest(excel_path)

    material_map = {}
    # 按阶段分组的报告
    report_by_stage = {}

    for item in manifest:
        stage = item['stage']
        doc_name = item['doc_name']
        storage_path = item['storage_path']
        aliases = item.get('aliases', [])
        path_exists_only = item.get('path_exists_only', False)

        key = (stage, doc_name)
        result = {
            'exists': False,
            'file_key': '',
            'filepath': '',
            'storage_path': storage_path,
            'alias': item.get('alias', ''),
            'path_exists_only': path_exists_only,
        }

        if storage_path:
            normalized = _normalize_path_for_match(storage_path)
            # 找到路径下的所有文件
            files_in_path = [
                info for info in metadata_infos
                if _path_matches(info.get('relative_path', ''), normalized)
            ]
            if path_exists_only:
                # 该字段为是：只要存放路径下存在文件即认定该材料存在
                if files_in_path:
                    matched = None
                    for info in files_in_path:
                        if _matches_doc(info.get('filename', ''), doc_name, aliases):
                            matched = info
                            break
                    if matched is None:
                        matched = files_in_path[0]
                    result['exists'] = True
                    result['file_key'] = matched.get('file_key', matched.get('filename', ''))
                    result['filepath'] = matched.get('filepath', '')
            elif len(files_in_path) == 1:
                info = files_in_path[0]
                result['exists'] = True
                result['file_key'] = info.get('file_key', info.get('filename', ''))
                result['filepath'] = info.get('filepath', '')
            elif len(files_in_path) > 1:
                # 多个文件，按文档名称或别名匹配
                matched = None
                for info in files_in_path:
                    if _matches_doc(info.get('filename', ''), doc_name, aliases):
                        matched = info
                        break
                if matched:
                    result['exists'] = True
                    result['file_key'] = matched.get('file_key', matched.get('filename', ''))
                    result['filepath'] = matched.get('filepath', '')
                # else: 多个文件但名称都不匹配，视为缺失
            # else: 路径下没有文件，视为缺失
        else:
            # 存放路径为 None，全局搜索文件名匹配
            matched = None
            for info in metadata_infos:
                if _matches_doc(info.get('filename', ''), doc_name, aliases):
                    matched = info
                    break
            if matched:
                result['exists'] = True
                result['file_key'] = matched.get('file_key', matched.get('filename', ''))
                result['filepath'] = matched.get('filepath', '')

        material_map[key] = result

        # 构建报告
        if stage not in report_by_stage:
            report_by_stage[stage] = []
        report_by_stage[stage].append({
            'doc_name': doc_name,
            'storage_path': storage_path or '（未指定路径，全局搜索）',
            'alias': item.get('alias', ''),
            'path_exists_only': path_exists_only,
            'exists': result['exists'],
            'file_key': result['file_key'],
        })

    return material_map, report_by_stage


def find_doc_in_manifest(doc_name, stage, manifest):
    """在清单中查找指定文档条目（支持名称/别名模糊匹配，跨阶段兜底）。

    先在指定阶段内查找，找不到再在所有阶段中查找（处理跨阶段材料引用）。
    返回匹配的清单条目（dict），未找到返回 None。
    """
    if not doc_name or not manifest:
        return None
    clean = re.sub(r'[（(][^）)]*[）)]', '', doc_name).strip()
    if not clean:
        return None

    def _search(candidates):
        # 精确匹配文档名称
        for m in candidates:
            if m['doc_name'] == clean:
                return m
        # 文档名称双向包含匹配
        for m in candidates:
            if clean in m['doc_name'] or m['doc_name'] in clean:
                return m
        # 别名匹配（检查要点里的叫法与清单别名对齐）
        for m in candidates:
            for alias in m.get('aliases', []):
                if clean == alias or clean in alias or alias in clean:
                    return m
        return None

    # 1. 先在当前阶段查找
    stage_candidates = [m for m in manifest if m['stage'] == stage]
    found = _search(stage_candidates)
    if found:
        return found

    # 2. 跨阶段兜底查找（如项目采购检查项引用项目前期的可研批复文件）
    found = _search(manifest)
    if found:
        return found

    return None
