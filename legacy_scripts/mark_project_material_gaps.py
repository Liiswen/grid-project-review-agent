# -*- coding: utf-8 -*-
"""根据真实项目资料核对结果，为确实没有对应材料的细则加醒目标记。"""

from pathlib import Path
import re


BASE = Path(__file__).resolve().parent.parent
RULE_DIR = BASE / "检查细则"
MARK = "【实际项目材料未发现】"

# 仅标记已核对确认没有对应材料、且该材料会影响该条核验的规则。
# 文件名不同但已有替代材料的规则不在此列。
MISSING = {
    "项目前期.txt": {2, 4, 7, 8, 9, 12, 13, 14, 15, 16, 17},
    "项目采购.txt": {4, 16, 18, 20, 21, 22, 23},
    "项目建设.txt": {20},
    "项目验收.txt": {4, 7},
    "决算.txt": {5, 6},
    "上下线.txt": set(),
    "转资.txt": set(),
}


def mark_file(path, sequences):
    text = path.read_text(encoding="utf-8")
    for seq in sorted(sequences, reverse=True):
        pattern = re.compile(rf"(?m)^【检查项{seq}】(?!" + re.escape(MARK) + r")")
        text, count = pattern.subn(f"【检查项{seq}】{MARK}", text, count=1)
        if count != 1:
            raise ValueError(f"{path.name} 未找到唯一检查项{seq}")
    path.write_text(text, encoding="utf-8")


def main():
    for filename, sequences in MISSING.items():
        mark_file(RULE_DIR / filename, sequences)
        print(filename, sorted(sequences))


if __name__ == "__main__":
    main()
