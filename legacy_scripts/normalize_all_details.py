# -*- coding: utf-8 -*-
"""统一七个阶段检查细则的字段、依据和操作指导结构。"""

from pathlib import Path
import re

BASE = Path(__file__).resolve().parent.parent
RULE_DIR = BASE / "检查细则"
POINT_DIR = BASE / "检查要点"
STAGES = ["项目前期", "项目采购", "项目建设", "上下线", "项目验收", "决算", "转资"]


def source_basis(stage, seq):
    point_file = POINT_DIR / f"{stage}.txt"
    if not point_file.exists():
        return ""
    text = point_file.read_text(encoding="utf-8")
    match = re.search(
        rf"【检查项{seq}】[^\n]+\n.*?【依据】([^\n]+)", text, re.S
    )
    if not match:
        return ""
    value = match.group(1).strip()
    return "" if value in {"未明确", "/"} else value


def parse_blocks(text):
    pattern = re.compile(
        r"^【检查项(?P<seq>\d+)】(?P<title>[^\n]+)\n(?P<body>.*?)(?=^【检查项\d+】|\Z)",
        re.M | re.S,
    )
    blocks = []
    for match in pattern.finditer(text):
        body = match.group("body")
        doc = re.search(r"(?m)^【对应文件】(?P<value>[^\n]+)", body)
        basis = re.search(r"(?m)^【依据】(?P<value>[^\n]+)", body)
        guide = re.search(r"(?ms)^操作指导：\s*(?P<value>.*)$", body)
        if not doc:
            continue
        blocks.append({
            "seq": int(match.group("seq")),
            "title": match.group("title").strip(),
            "documents": doc.group("value").strip(),
            "basis": basis.group("value").strip() if basis else "",
            "guidance": guide.group("value").strip() if guide else "",
        })
    return blocks


def clean_guidance(text):
    # 旧版细则的 guidance 可能已经包含合格/不合格/人工审核段落；只保留步骤正文。
    text = re.split(r"(?m)^【(?:合格标准|不合格标准|无法判断及人工审核建议)】", text)[0]
    text = re.sub(r"\n={5,}\n", "\n", text)
    text = re.sub(r"(?m)^判定边界：.*$", "", text)
    text = re.sub(r"(?m)^人工审核建议：.*$", "", text)
    # “原始检查要求”已经保留在【检查项】本身，步骤中不再重复抄写。
    text = re.sub(r"(?m)^原始检查要求：.*(?:\n|$)", "", text)
    # 这些内容由 render() 统一输出到对应字段，不能留在步骤中。
    text = re.sub(r"(?m)^已取得对应材料，且材料中的项目名称.*$", "", text)
    text = re.sub(r"(?m)^已取得并实际读取相关材料，能够明确证明.*$", "", text)
    text = re.sub(r"(?m)^材料未提供、关键页面未展示.*$", "", text)
    text = text.replace("无法确定", "无法判断")
    # 删除误混入步骤的章节标题和规则编写说明。
    text = re.sub(r"(?m)^\s*(?:步骤\s*\d+：)?[一二三四五六七八九十]+、.*检查要点\s*$", "", text)
    text = re.sub(r"(?m)^\s*(?:步骤\s*\d+：)?每个检查项应输出以下内容：\s*$", "", text)
    # “审查输出要求”属于整个平台的输出规范，不是某一检查项的核验步骤。
    text = re.split(r"(?m)^\s*步骤\s*\d+：.*审查输出要求\s*$", text)[0]
    # 这是旧运行结果写入的项目个案结论，不属于通用细则，不能混入规则。
    text = re.sub(r"(?ms)^现有已检索资料中缺少.*?(?=\n\s*\n|\Z)", "", text)
    # 允许脚本重复执行：去掉上一轮生成的标准标签，避免标签嵌套。
    for label in ("核验对象", "核验步骤", "合格标准", "不合格标准", "无法判断及人工审核建议"):
        text = re.sub(rf"(?m)^【{label}】\s*", "", text)
    # 清除重复执行时遗留的通用核验对象句。
    lines = text.splitlines()
    result = []
    previous = None
    step_number = 1
    for line in lines:
        normalized = re.sub(r"\s+", "", line)
        if normalized.startswith("围绕‘") and normalized.endswith("核对对应文件及其关联材料。"):
            continue
        if line.strip() and line.strip() == previous:
            continue
        numbered = re.match(r"^\s*步骤\s*\d+\s*[：:、.]\s*(.*)$", line)
        if numbered:
            line = f"步骤{step_number}：{numbered.group(1).strip()}"
            step_number += 1
        elif line.strip() and not line.startswith((" ", "\t")):
            line = f"步骤{step_number}：{line.strip()}"
            step_number += 1
        result.append(line)
        if line.strip():
            previous = line.strip()
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    text = "\n".join(result).strip()
    # 将过细的步骤压缩为“定位材料—核对关系—处理例外/判定”三步，保留全部原文。
    step_matches = list(re.finditer(r"(?ms)^步骤(\d+)：(.*?)(?=^步骤\d+：|\Z)", text))
    if len(step_matches) > 3:
        compact = []
        for index, match in enumerate(step_matches[:2], 1):
            compact.append(f"步骤{index}：{match.group(2).strip()}")
        tail = "；".join(match.group(2).strip() for match in step_matches[2:])
        compact.append(f"步骤3：{tail}")
        text = "\n".join(compact)
    return text or "根据对应文件的正文、签章页、日期、金额及相关附件进行核验。"


def clean_basis(value):
    value = re.sub(r"(?:规章制度依据：\s*)+", "规章制度依据：", value)
    value = re.sub(r"^规章制度依据：\s*", "", value)
    value = re.sub(r"^附件2对应检查要点及相关制度条款\s*", "", value)
    value = re.sub(r"(?:；?\s*规章制度依据：)?附件2《电网数字化项目检查要点.xlsx》开发实施项目检查表（阶段：[^）]+）", "", value)
    value = re.sub(r"(?:；?\s*附件2未列明具体制度名称，按相关制度条款核验)", "", value)
    value = re.sub(r"(?:；\s*){2,}", "；", value).strip("； ")
    while "规章制度依据：规章制度依据：" in value:
        value = value.replace("规章制度依据：规章制度依据：", "规章制度依据：")
    return "" if value in {"", "未明确", "/"} else value


def render(stage, blocks):
    lines = [
        "国家电网有限公司电网数字化项目检查细则",
        f"阶段：{stage}",
        "来源：附件2《电网数字化项目检查要点.xlsx》开发实施项目检查表及D:\\桌面\\实习\\规章制度相关文件",
        "说明：每条均保留附件2对应阶段要点，并结合规章制度细化核验方法；标有‘实际项目材料未发现’的条目仅表示本次实际项目资料中未找到对应证据，不直接等同于不合格。",
        "",
    ]
    for block in blocks:
        original_basis = clean_basis(block["basis"])
        excel_basis = f"附件2《电网数字化项目检查要点.xlsx》开发实施项目检查表（阶段：{stage}）"
        # 项目前期、项目采购是早期扩展细则，序号不是附件2原始序号，不能按序号硬套依据。
        if stage not in {"项目前期", "项目采购"}:
            detail_basis = source_basis(stage, block["seq"])
        else:
            detail_basis = ""
        other_basis = detail_basis or ("" if original_basis in {"", "未明确", "/"} else original_basis)
        other_basis = other_basis.replace("规章制度依据：", "").strip("； ")
        if other_basis == "附件2对应检查要点及相关制度条款":
            other_basis = ""
        basis = excel_basis + (f"；规章制度依据：{other_basis}" if other_basis else "；附件2对应检查要点及相关制度条款")
        guide = clean_guidance(block["guidance"])
        documents = block["documents"]
        # 该条按项目类型适用不同试运行期限，不能误用“上线运行验收日期”作比较终点。
        if stage == "项目验收" and block["seq"] == 8:
            guide = (
                "步骤1：先按项目年度和新建/续建版本类型确定适用的试运行期限，再提取上线日期和上线试运行验收申请日期，"
                "按自然日期计算：2024年及以前项目上线时间应大于3个月；2025年及以后新建系统上线至验收申请应大于3个月；"
                "续建系统大版本应大于1个月；小版本上线即进入正式运行。项目类型或日期无法确认时列为无法判断。"
            )
        if stage == "上下线" and block["seq"] == 8:
            documents = "项目上线运行结果确认单、应用成效确认单"
            guide = (
                "步骤1：提取项目上线运行结果确认单出具日期和应用成效确认单出具日期，确认应用成效确认单应在前者出具后一个月内完成，"
                "并核对应用成效确认单中的应用成效、业务部门确认和签字盖章；任一日期或关键签章无法确认时列为无法判断。"
            )
        if stage == "项目验收" and block["seq"] == 8:
            documents = "上线试运行验收申请单、上线试运行验收意见、验收申请单"
        display_title = block["title"].replace("【实际项目材料未发现】", "").strip()
        lines.extend([
            f"【检查项{block['seq']}】{block['title']}",
            f"【对应文件】{documents}",
            f"【依据】{basis}",
            "操作指导：",
            f"【核验对象】围绕‘{display_title}’核对对应文件及其关联材料。",
            "【核验步骤】",
            guide,
            "【合格标准】已取得对应材料，且材料中的项目名称、版本、日期、金额、签字盖章和正文内容（按本条适用项）能够证明检查要求全部满足。",
            "【不合格标准】已取得并实际读取相关材料，能够明确证明存在与本条检查要求相反的事实；问题描述必须写出具体文件、字段、页码或日期关系。",
            "【无法判断及人工审核建议】材料未提供、关键页面未展示、文件虽有但字段无法辨认、系统记录未提供或关联材料无法交叉验证时，结论为‘无法判断’，不得仅因文件名不同判定缺失；建议人工核验本条对应文件的原件/完整页、具体字段和系统记录。",
            "",
        ])
    return "\n".join(lines).rstrip() + "\n"


def main():
    for stage in STAGES:
        path = RULE_DIR / f"{stage}.txt"
        blocks = parse_blocks(path.read_text(encoding="utf-8"))
        if not blocks:
            raise ValueError(f"未解析到{path}")
        path.write_text(render(stage, blocks), encoding="utf-8")
        print(stage, len(blocks))


if __name__ == "__main__":
    main()
