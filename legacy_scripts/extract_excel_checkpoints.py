# -*- coding: utf-8 -*-
"""从附件2工作簿生成按阶段归档的检查要点 TXT。"""

from collections import defaultdict
from pathlib import Path
import re
import sys

from openpyxl import load_workbook


PROJECT_TYPES = {
    "开发实施项目检查表": "开发实施项目",
    "软硬件购置项目检查表": "软硬件购置项目",
    "数据工程项目": "数据工程项目",
    "咨询设计类项目": "咨询设计类项目",
    "数字化专业业务运营类项目": "数字化专业业务运营类项目",
}

# 项目档案采用统一的九类目录归档。该映射仅用于快速定位证据，不增加或改变附件2检查要求。
PATH_HINTS = {
    "可行性研究报告": ("01项目可研/01可研报告",),
    "可研评审意见": ("01项目可研/02可研评审意见",),
    "可研批复文件": ("01项目可研/03可研批复",),
    "采购文件及采购依据": ("02采购资料",),
    "项目合同": ("03项目合同/01合同",),
    "合同补充协议": ("03项目合同/01合同",),
    "分包规范": ("06实施文档/11项目分包管理",),
    "公司项目启动或推进文件": ("04启动及需求文档/01项目启动资料",),
    "需求规格说明书": ("04启动及需求文档/02需求说明书",),
    "系统概要设计说明书": ("05设计开发文档/01系统设计报告",),
    "详细设计说明书": ("05设计开发文档/01系统设计报告",),
    "数据字典": ("05设计开发文档/02数据字典",),
    "集成测试报告": ("05设计开发文档/03集成测试",),
    "出厂测试报告": ("05设计开发文档/04出厂测试",),
    "第三方测试报告": ("05设计开发文档/05第三方测试",),
    "用户确认测试报告": ("05设计开发文档/06用户确认测试",),
    "实施方案": ("06实施文档/01实施方案",),
    "安全防护方案": ("06实施文档/02安全防护",),
    "用户手册": ("06实施文档/03用户手册",),
    "系统部署方案": ("06实施文档/04系统部署",),
    "系统管理员手册": ("06实施文档/05系统管理员手册",),
    "系统应急预案": ("06实施文档/06系统应急预案",),
    "上线试运行申请单": ("06实施文档/07上线试运行及试运行验收/01上线试运行申请单",),
    "上线试运行验收申请单": ("06实施文档/07上线试运行及试运行验收/04上线试运行验收申请单",),
    "上线试运行验收意见": ("06实施文档/07上线试运行及试运行验收/05上线试运行验收意见",),
    "项目上线运行结果确认单": ("06实施文档/07上线试运行及试运行验收/06项目上线运行结果确认单",),
    "验收申请单": ("07验收文档/01验收申请单",),
    "用户报告": ("07验收文档/02用户报告",),
    "竣工总结报告": ("07验收文档/03竣工总结报告",),
    "验收意见": ("07验收文档/04验收意见",),
    "项目建设内容完成确认单": ("07验收文档/06项目建设内容完成确认单",),
    "应用成效确认单": ("07验收文档/05应用成效确认单",),
    "项目结算报告": ("08结算转资审计/01项目结算报告",),
    "项目决算报告": ("08结算转资审计/02项目决算报告",),
    "转资": ("08结算转资审计/03转资材料",),
    "决算审计报告": ("08结算转资审计/04项目审计报告",),
}


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def find_column(headers, keywords):
    for index, header in enumerate(headers, 1):
        text = clean(header)
        if any(keyword in text for keyword in keywords):
            return index
    return None


def extract(workbook_path):
    workbook = load_workbook(workbook_path, data_only=True)
    by_stage = defaultdict(lambda: defaultdict(list))

    for sheet in workbook.worksheets:
        # 当前平台仅审查开发实施项目，附件中的其他项目类型不写入规则库。
        if sheet.title != "开发实施项目检查表":
            continue
        project_type = PROJECT_TYPES[sheet.title]
        headers = [sheet.cell(2, column).value for column in range(1, sheet.max_column + 1)]
        sequence_col = find_column(headers, ("序号",))
        stage_col = find_column(headers, ("项目阶段",))
        link_col = find_column(headers, ("项目环节",))
        document_col = find_column(headers, ("需要上传",))
        point_col = find_column(headers, ("上传前检查", "上传前检查事项"))
        basis_col = find_column(headers, ("依据文件",))
        requirement_col = find_column(headers, ("相关要求",))
        remark_col = find_column(headers, ("备注",))

        if not all((stage_col, point_col)):
            continue

        inherited = {"stage": "", "link": "", "document": ""}
        for row in range(3, sheet.max_row + 1):
            values = [sheet.cell(row, column).value for column in range(1, sheet.max_column + 1)]
            if not any(value is not None for value in values):
                continue
            # 说明和模板位于表末，不是检查项。
            first = clean(values[0])
            if first.startswith("备注：") or "情况说明（模板）" in first:
                break

            stage = clean(sheet.cell(row, stage_col).value)
            link = clean(sheet.cell(row, link_col).value) if link_col else ""
            document = clean(sheet.cell(row, document_col).value) if document_col else ""
            if stage:
                inherited["stage"] = stage
            if link:
                inherited["link"] = link
            if document:
                inherited["document"] = document

            point = clean(sheet.cell(row, point_col).value)
            if not point or not inherited["stage"]:
                continue

            item = {
                "excel_seq": clean(sheet.cell(row, sequence_col).value) if sequence_col else "",
                "link": inherited["link"],
                "document": inherited["document"],
                "point": point,
                "basis": clean(sheet.cell(row, basis_col).value) if basis_col else "",
                "requirement": clean(sheet.cell(row, requirement_col).value) if requirement_col else "",
                "remark": clean(sheet.cell(row, remark_col).value) if remark_col else "",
                "source": f"{sheet.title}第{row}行",
            }
            by_stage[inherited["stage"]][project_type].append(item)
    return by_stage


def render_stage(stage, project_types):
    lines = [
        "国家电网有限公司电网数字化项目检查要点",
        "来源：附件2.电网数字化项目检查要点.xlsx",
        f"阶段：{stage}",
        "说明：以下内容逐条转录自附件2；按项目类型适用，未填写的字段不作补充推定。",
        "",
    ]
    for project_type, items in project_types.items():
        lines.extend([f"【项目类型】{project_type}", ""])
        for index, item in enumerate(items, 1):
            path_hints = []
            for label, paths in PATH_HINTS.items():
                if label in item['document']:
                    path_hints.extend(paths)
            lines.extend([
                f"【检查项{index}】{item['point']}",
                f"【Excel序号】{item['excel_seq'] or '未填写'}",
                f"【项目环节】{item['link'] or '未填写'}",
                f"【对应文件】{item['document'] or '未明确'}",
                f"【依据】{item['basis'] or '未明确'}",
                f"【相关要求】{item['requirement'] or '未明确'}",
                f"【备注】{item['remark'] or '无'}",
                f"【建议检索路径】{'；'.join(dict.fromkeys(path_hints)) or '未明确'}",
                f"【来源定位】{item['source']}",
                "",
            ])
    return "\n".join(lines).rstrip() + "\n"


def main():
    workbook = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    output_dir.mkdir(parents=True, exist_ok=True)
    for old_file in output_dir.glob("*.txt"):
        old_file.unlink()
    by_stage = extract(workbook)
    for stage, project_types in by_stage.items():
        (output_dir / f"{stage}.txt").write_text(render_stage(stage, project_types), encoding="utf-8")
    for stage, project_types in by_stage.items():
        print(f"{stage}: " + ", ".join(f"{name}{len(items)}项" for name, items in project_types.items()))


if __name__ == "__main__":
    main()
