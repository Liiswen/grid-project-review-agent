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

# 一条细则往往要求交叉查看多类主证据。这里将细则中的材料名称映射为可在
# 文件名、目录名或文件类型中识别的别名，保证候选集至少覆盖每一类已上传主证据。
DOCUMENT_ALIASES = {
    '项目可行性研究报告': ('可行性研究', '可研报告'),
    '可研评审意见': ('可研评审',),
    '可研批复文件': ('可研批复', '可研批复文件'),
    '公司下达计划文件': ('计划下达', '计划文件', '项目清单'),
    '各单位下达计划文件': ('计划下达', '计划文件', '项目清单'),
    '中标/成交通知书': ('中标通知书', '成交通知书'),
    '项目合同': ('实施服务合同', '软件开发合同', '项目合同'),
    '招标文件/谈判文件/询价文件': ('招标文件', '谈判文件', '询价文件'),
    '采购技术规范书': ('技术规范书',),
    '投标报价文件': ('报价文件', '报价明细表', '价格投标文件'),
    '合同流转单': ('合同审批单', '合同流转单'),
    '分包审批': ('分包审批', '分包备案', '分包情况'),
    '开工通知': ('开工通知', '启动会'),
}


# 这不是项目事实库，而是每类审查所必需的最小交叉证据集合。规则文件写的是
# “检查什么、到哪个标准目录找”；这里仅将抽象材料名稳定地映射回已上传文件，
# 防止“分包规范”“采购文件及采购依据”这类宽泛对应文件名把关键证据筛掉。
SEMANTIC_EVIDENCE_GROUPS = (
    (
        ('分包',),
        ('分包备案', '分包申请', '分包审批', '分包合同', '分包商', '分包情况', '不分包承诺'),
        6,
    ),
    (
        ('中标金额', '可研批复金额', '合同金额', '不低于', '金额占'),
        # 金额核验的前五份视觉候选必须覆盖“批复、通知书、全部主合同”；
        # 报价材料作为补充，不能挤掉任一拆分合同。
        ('可研批复', '中标通知书', '成交通知书', '实施服务合同', '软件开发合同', '价格投标', '报价明细表'),
        6,
    ),
    (
        ('中标通知书下达', '30日内', '签订日期', '经法系统', '经法日期'),
        # 时限比对先传通知书和全部盖章主合同；审批/流转单用于随后核验经法日期，
        # 不得占用拆分合同的视觉候选位置。
        ('中标通知书', '成交通知书', '实施服务合同', '软件开发合同', '合同审批单', '合同流转单'),
        6,
    ),
    (
        ('乙方名称', '供应商名称', '公司更名'),
        ('中标通知书', '成交通知书', '商务投标文件', '投标文件', '实施服务合同', '软件开发合同'),
        6,
    ),
)


def select_candidates(checkpoint, file_infos, primary_file='', limit=8, with_details=False):
    """根据文件名、目录名、首页标题和检查项语义返回有序候选文件名。

    with_details=True 时同时返回每个候选的打分及入选通道，供流程日志打印；
    打分只用于说明候选为何入选，不改变既有的候选排序结果。
    """
    query = ' '.join([
        checkpoint.get('doc_name', ''), checkpoint.get('check_point', ''),
        checkpoint.get('requirement', ''), checkpoint.get('stage', ''),
    ])
    terms = _terms(query)
    path_hints = [_normalise_path_hint(hint) for hint in checkpoint.get('search_paths', [])]
    path_hints = [hint for hint in path_hints if hint]
    ranked = []

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
    for info in source_infos:
        score = base_score(info)
        if score:
            ranked.append((score, info.get('file_key', info.get('filename', ''))))
    ranked.sort(key=lambda item: (-item[0], item[1]))

    result = []
    details = {}

    def add(filename, score, source):
        """记录候选的入选通道；同一文件保留首次入选时的打分与来源。"""
        if filename in result:
            return
        result.append(filename)
        details[filename] = {'score': score, 'source': source}

    primary_info = next(
        (info for info in file_infos if info.get('file_key', info.get('filename')) == primary_file),
        None,
    )
    if primary_info:
        # 主匹配可能来自建议检索路径之外的目录，按其自身排序分展示即可。
        add(primary_file, base_score(primary_info), '上一轮主匹配')

    # 先放入检查语义必需的交叉证据。每个文件仍须真实存在于本次上传资料中；
    # 不依据项目名称、金额、日期或固定文件名作任何推定。
    for triggers, aliases, take_count in SEMANTIC_EVIDENCE_GROUPS:
        if not any(trigger in query for trigger in triggers):
            continue
        matches = _match_aliases(file_infos, aliases)
        for score, preferred in matches[:take_count]:
            add(preferred, score, '语义证据组')

    # 对细则明确列出的每种既有主材料至少保留一个最相关文件。这样“可研+评审+批复”
    # 等交叉核验不会退化为只读其中一份，也不会因模型首轮匹配失误把已上传材料排除。
    # “对应文件”可能写得很泛（如“采购文件及采购依据”），具体材料类型常写在
    # “相关要求”中；两处都要参与别名匹配，不能只看对应文件字段。
    documents = ' '.join((checkpoint.get('doc_name', ''), checkpoint.get('requirement', '')))
    for label, aliases in DOCUMENT_ALIASES.items():
        if label not in documents:
            continue
        matches = []
        for info in file_infos:
            filename_path = ' '.join((info.get('filename', ''), info.get('relative_path', '')))
            score = sum(20 for alias in aliases if alias in filename_path)
            if score:
                matches.append((score, info.get('file_key', info.get('filename', ''))))
        if matches:
            ordered = sorted(matches, key=lambda item: (
                -item[0], '09支付文件' in item[1], '03项目合同/01合同' not in item[1].replace('\\', '/'),
                '__nested__' not in item[1] and '02采购资料' not in item[1],
                len(item[1]), item[1],
            ))
            take_count = 2 if label == '项目合同' else 1
            for score, preferred in ordered[:take_count]:
                add(preferred, score, '材料别名')
    for score, filename in ranked:
        add(filename, score, '关键词/路径打分')
        if len(result) >= limit:
            break
    if with_details:
        return result, details
    return result


def _match_aliases(file_infos, aliases):
    """按文件名和路径返回真实匹配文件及其打分（分数越高越优先）。"""
    matches = []
    for info in file_infos:
        filename = info.get('filename', '')
        relative_path = info.get('relative_path', '')
        # “中标通知书”目录下常同时有报价表；文件名命中必须高于仅目录命中，
        # 否则会把报价表误作为通知书主证据。
        score = 0
        # aliases 的排列就是该检查类型的取证顺序（例如先批复、再通知书、再报价）。
        # 用足够大的权重避免同一目录下的“报价表”抢在“中标通知书”之前。
        for index, alias in enumerate(aliases):
            priority = len(aliases) - index
            if alias in filename:
                score += priority * 100
            if alias in relative_path:
                score += priority * 10
        if score:
            matches.append((score, info.get('file_key', info.get('filename', ''))))
    return sorted(matches, key=lambda item: (
        -item[0], '09支付文件' in item[1],
        '03项目合同/01合同' not in item[1].replace('\\', '/'),
        '__nested__' not in item[1] and '02采购资料' not in item[1],
        len(item[1]), item[1],
    ))


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
