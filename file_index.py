# -*- coding: utf-8 -*-
"""轻量文件索引与候选检索：用于匹配，不做全文解析或图像转换。"""

import re


# 文件名、目录名和首页摘要中高频出现的业务实体。先命中这些完整词，
# 再以短词补充，避免把一整段中文操作说明当成一个永远匹配不到的词。
DOMAIN_TERMS = (
    '可行性研究报告', '可研报告', '可研评审意见', '可研批复', '项目说明书',
    '计划下达', '储备库', '项目储备', '中标通知书', '成交通知书', '招标公告',
    '招标文件', '技术规范书', '投标文件', '报价明细表', '合同审批单',
    '软件开发合同', '实施服务合同', '补充协议', '合同变更', '合同流转单',
    '营业执照', '授权委托书', '授权书', '分包备案', '分包情况说明',
    '分包合同', '分包审批', '不分包承诺', '考勤记录', '开工通知',
    '启动会', '进度计划', 'IRS', '经法系统', '软件著作权', '查重记录',
)


def select_candidates(checkpoint, file_infos, limit=8):
    """按关键词与路径得分返回前 limit 个候选文件名（降序）。"""
    query = ' '.join([
        checkpoint.get('doc_name', ''), checkpoint.get('check_point', ''),
        checkpoint.get('requirement', ''), checkpoint.get('stage', ''),
    ])
    terms = _terms(query)
    path_hints = [_normalise_path_hint(hint) for hint in checkpoint.get('search_paths', [])]
    path_hints = [hint for hint in path_hints if hint]

    def base_score(info):
        """路径命中、关键词命中与文件类型匹配的排序分。"""
        filename = info.get('filename', '')
        path = info.get('relative_path', '')
        normalized_path = path.replace('\\', '/').lower()
        score = 0
        # 固定归档目录命中优先级高于模糊语义，适配规范化项目资料包。
        if any(hint in normalized_path for hint in path_hints):
            score += 100
        for term in terms:
            if term in filename:
                score += 8
            if term in path:
                score += 4
        file_type = info.get('file_type', '')
        if file_type and file_type in query:
            score += 6
        return score

    # 命中固定归档路径时，仅在该目录内择优，防止“可研目录”等相近资料抢占主证据。
    hinted_infos = [
        info for info in file_infos
        if any(hint in info.get('relative_path', '').replace('\\', '/').lower() for hint in path_hints)
    ]
    source_infos = hinted_infos or file_infos

    ranked = []
    for info in source_infos:
        score = base_score(info)
        if score:
            ranked.append((score, info.get('file_key', info.get('filename', ''))))
    ranked.sort(key=lambda item: (-item[0], item[1]))
    return [filename for _, filename in ranked[:limit]]


def _normalise_path_hint(hint):
    """将检查要点中的说明性括号从真实目录路径中剥离。

    例如“02采购资料/02招标文件（含全部内嵌ZIP）”应匹配实际目录
    “02采购资料/02招标文件”，括号说明仍保留在规则文本中供模型执行。
    """
    value = hint.replace('\\', '/').strip().lower()
    value = re.sub(r'[（(][^）)]*[）)]', '', value)
    return value.rstrip('/').strip()


def _terms(text):
    stop_words = {'项目', '文件', '材料', '检查', '内容', '要求', '相关', '应当', '是否', '以及', '开发', '实施', '对应文件'}
    terms = []

    def add(term):
        term = term.strip()
        if len(term) >= 2 and term not in stop_words and term not in terms:
            terms.append(term)

    for term in DOMAIN_TERMS:
        if term in text:
            add(term)

    # 中文正则会把连续的一整段句子视为一个词。将其拆为有辨识度的短语，
    # 使“合同签署页”“技术规范书”等可在文件名和目录中实际命中。
    for span in re.findall(r'[\u4e00-\u9fff]{2,}', text):
        if len(span) <= 8:
            add(span)
        else:
            for size in (6, 5, 4, 3, 2):
                for index in range(len(span) - size + 1):
                    add(span[index:index + size])
                    if len(terms) >= 80:
                        return terms
    for term in re.findall(r'[A-Za-z]{2,}|\d+%', text):
        add(term)
    return terms[:80]
