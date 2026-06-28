#!/usr/bin/env python3
"""푸른조경 조경 XLS — 「1). 식재」 블록만 추출 → 마크다운."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import xlrd

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SRC = ROOT / "05_내역서" / "01_260620화성시마도면총원리청원지구 조경.XLS"
DEFAULT_OUT = ROOT / "05_내역서" / "내역서작업" / "01_260620_푸른조경_식재내역.md"

ITEM_RE = re.compile(r"^[가-힣]\)\.\s")


def extract(src: Path = DEFAULT_SRC, out: Path = DEFAULT_OUT) -> int:
    wb = xlrd.open_workbook(str(src))
    sh = wb.sheet_by_name("내역서")
    lines = [
        "# 푸른조경 식재 내역 추출",
        "",
        f"- **원본**: `{src.name}`",
        "- **발행·기준일**: 2026. 6. 20.",
        "- **범위**: 「1). 식재」 블록(소공원·경관녹지·가로수)",
        "",
        "| 구역 | 품명 | 규격 | 수량 | 단위 |",
        "|---|---|---:|---:|---|",
    ]
    zone = ""
    in_plant = False
    n = 0
    for r in range(sh.nrows):
        c0 = str(sh.cell_value(r, 0)).strip()
        if not c0:
            continue
        if c0.startswith("1).") and "식재" in c0:
            in_plant = True
            continue
        if c0.startswith(("2).", "3).", "4).")):
            in_plant = False
            continue
        if re.match(r"^[가-힣]\.\s", c0):
            zone = c0
            continue
        if not in_plant or not ITEM_RE.match(c0):
            continue
        qty = sh.cell_value(r, 2)
        unit = str(sh.cell_value(r, 3)).strip()
        spec = str(sh.cell_value(r, 1)).strip()
        if not (isinstance(qty, (int, float)) and qty > 0 and unit and unit != "식"):
            continue
        name = ITEM_RE.sub("", c0).strip()
        lines.append(f"| {zone} | {name} | {spec} | {qty:g} | {unit} |")
        n += 1
    lines.extend(["", f"총 **{n}**행.", "", "끝.", ""])
    out.write_text("\n".join(lines), encoding="utf-8")
    return n


if __name__ == "__main__":
    n = extract()
    print(f"식재 {n}행 → {DEFAULT_OUT}")
