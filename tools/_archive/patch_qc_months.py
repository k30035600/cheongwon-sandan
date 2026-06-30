#!/usr/bin/env python3
"""품질관리자인건비 투입 — 개월 수 일괄 변경(공내역서 XLS)."""
from __future__ import annotations

import sys
from pathlib import Path

import xlrd
import xlwt

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
TARGET_MONTHS = 14
OLD_MONTHS = 34

FILES = [
    ROOT / "05_내역서" / "공내역서" / "01_화성 청원지구 토목.XLS",
    ROOT / "05_내역서" / "공내역서" / "01_화성 청원지구 토목(조경)_통합원본.XLS",
]


def copy_xls(src: Path, dst: Path, patches: list[tuple[int, int, int, float]]) -> int:
    rb = xlrd.open_workbook(str(src))
    wb = xlwt.Workbook()
    n = 0
    patch_map = {(r, c): v for _, r, c, v in patches}
    for si in range(rb.nsheets):
        rs = rb.sheet_by_index(si)
        ws = wb.add_sheet(rs.name)
        for r in range(rs.nrows):
            for c in range(rs.ncols):
                if (r, c) in patch_map:
                    ws.write(r, c, patch_map[(r, c)])
                    n += 1
                else:
                    ct = rs.cell_type(r, c)
                    val = rs.cell_value(r, c)
                    if ct == xlrd.XL_CELL_NUMBER:
                        ws.write(r, c, val)
                    elif ct == xlrd.XL_CELL_DATE:
                        ws.write(r, c, val)
                    else:
                        ws.write(r, c, str(val) if val is not None else "")
    wb.save(str(dst))
    return n


def find_qc_rows(path: Path) -> list[tuple[int, int]]:
    rb = xlrd.open_workbook(str(path))
    sh = rb.sheet_by_name("내역서")
    hits: list[tuple[int, int]] = []
    for r in range(sh.nrows):
        name = str(sh.cell_value(r, 0)).strip()
        unit = str(sh.cell_value(r, 3)).strip()
        qty = sh.cell_value(r, 2) if sh.cell_type(r, 2) == xlrd.XL_CELL_NUMBER else None
        if "품질관리자인건비" in name and unit == "개월" and qty == OLD_MONTHS:
            hits.append((r, 2))
    return hits


def main() -> None:
    total = 0
    for path in FILES:
        if not path.exists():
            print(f"[건너뜀] 없음: {path}")
            continue
        rows = find_qc_rows(path)
        if not rows:
            print(f"[없음] {path.name} — {OLD_MONTHS}개월 행 없음")
            continue
        patches = [(path.name, r, c, float(TARGET_MONTHS)) for r, c in rows]
        tmp = path.with_suffix(".patched.xls")
        copy_xls(path, tmp, patches)
        tmp.replace(path)
        total += len(rows)
        print(f"갱신: {path.name} — 품질관리자인건비 {OLD_MONTHS}→{TARGET_MONTHS}개월 ({len(rows)}건)")
    print(f"완료 {total}건")


if __name__ == "__main__":
    main()
