# -*- coding: utf-8 -*-
"""将阶段检查要点 TXT 转换为供审查引擎使用的结构化检查项。"""

import os
import re


def load_rule_checkpoints(rule_path, stage_name):
    """读取阶段检查要点 TXT，返回结构化检查项。"""
    if not os.path.exists(rule_path):
        raise FileNotFoundError(f"检查要点文件不存在: {rule_path}")

    with open(rule_path, "r", encoding="utf-8") as rule_file:
        content = rule_file.read()

    excel_pattern = re.compile(
        r"【检查项(?P<seq>\d+)】(?P<title>[^\n]+)\n"
        r"【Excel序号】(?P<excel_seq>[^\n]*)\n"
        r"【项目环节】(?P<link>[^\n]*)\n"
        r"【对应文件】(?P<documents>[^\n]*)\n"
        r"【依据】(?P<basis>[^\n]*)\n"
        r"【相关要求】(?P<requirement>[^\n]*)\n"
        r"【备注】(?P<remark>[^\n]*)\n"
        r"【建议检索路径】(?P<search_paths>[^\n]*)\n"
        r"【来源定位】(?P<source>[^\n]*)",
        re.S,
    )
    checkpoints = []
    for match in excel_pattern.finditer(content):
        documents = match.group("documents").strip()
        requirement = match.group("requirement").strip()
        checkpoints.append({
            "seq": int(match.group("seq")),
            "stage": stage_name,
            "link": match.group("link").strip() if match.group("link").strip() != "未填写" else stage_name,
            "doc_name": "" if documents == "未明确" else documents,
            "check_point": match.group("title").strip(),
            "basis": "" if match.group("basis").strip() == "未明确" else match.group("basis").strip(),
            "requirement": requirement if requirement != "未明确" else match.group("title").strip(),
            "risk": "",
            "possible_gap_materials": "【可能缺口材料" in documents or stage_name == '项目采购',
            "search_paths": [] if match.group("search_paths").strip() == "未明确" else match.group("search_paths").strip().split('；'),
            # 默认运行的“检查要点”须自包含取证规则；不再从代码维护项目专属证据链。
            "evidence_hint": "",
            "stage_evidence_guide": "以本检查项的【相关要求】和【建议检索路径】为准；先检索路径内材料及内嵌压缩包，再交叉取证。关键证据未定位或无法辨认时列为无法判断，不得仅因候选文件未命中认定未提供。",
            "source": match.group("source").strip(),
        })
    if not checkpoints:
        raise ValueError(f"未能从{stage_name}检查要点中解析出检查项: {rule_path}")
    return checkpoints
