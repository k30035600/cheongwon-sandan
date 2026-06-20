#!/usr/bin/env python3
"""01_화성 청원지구 토목(조경).XLS → 토목·조경 2개 XLS 분리.

조경 블록: 원본 row 814~934 (두 번째 「10. 조 경 공」 및 수량 내역·안전관리비).
토목: row 4~813 + row 935~끝 (품질관리비·제경비 제외공종 등).
"""
from __future__ import annotations

import sys
from pathlib import Path

import xlrd
import xlwt

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "05_내역서" / "공내역서"
SRC = SRC_DIR / "01_화성 청원지구 토목(조경)_통합원본.XLS"
OUT_TOK = SRC_DIR / "01_화성 청원지구 토목.XLS"
OUT_JO = SRC_DIR / "01_화성 청원지구 조경.XLS"

# Excel 1-based row numbers (inclusive)
JO_START = 814
JO_END = 934


def row_indices_tok(nrows: int) -> list[int]:
    """0-based row indices for 토목 file."""
    rows = list(range(0, 3))  # header rows 1~3
    rows += list(range(3, JO_START - 1))  # rows 4~813
    rows += list(range(JO_END, nrows))  # rows 935~end
    return rows


def row_indices_jo(nrows: int) -> list[int]:
    """0-based row indices for 조경 file."""
    rows = list(range(0, 4))  # header + title row 4
    rows += list(range(JO_START - 1, JO_END))  # rows 814~934
    return rows


def copy_cell(src: xlrd.sheet.Sheet, r: int, c: int, dst: xlwt.Worksheet, out_r: int) -> None:
    cell = src.cell(r, c)
    ct = cell.ctype
    val = cell.value
    if ct == xlrd.XL_CELL_EMPTY:
        return
    if ct == xlrd.XL_CELL_TEXT:
        dst.write(out_r, c, val)
    elif ct == xlrd.XL_CELL_NUMBER:
        dst.write(out_r, c, val)
    elif ct == xlrd.XL_CELL_DATE:
        dst.write(out_r, c, val)
    elif ct == xlrd.XL_CELL_BOOLEAN:
        dst.write(out_r, c, val)
    elif ct == xlrd.XL_CELL_ERROR:
        dst.write(out_r, c, xlwt.Formula(f'ERROR({val})'))
    else:
        dst.write(out_r, c, val)


def write_filtered(src_sh: xlrd.sheet.Sheet, row_indices: list[int], out_path: Path) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(src_sh.name)
    for out_r, src_r in enumerate(row_indices):
        for c in range(src_sh.ncols):
            copy_cell(src_sh, src_r, c, ws, out_r)
    wb.save(str(out_path))


def copy_sheet_as_is(src_sh: xlrd.sheet.Sheet, out_path: Path, sheet_name: str | None = None) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    name = sheet_name or src_sh.name
    ws = wb.add_sheet(name)
    for r in range(src_sh.nrows):
        for c in range(src_sh.ncols):
            copy_cell(src_sh, r, c, ws, r)
    wb.save(str(out_path))


def split_workbook(src_path: Path, tok_path: Path, jo_path: Path) -> tuple[int, int]:
    rb = xlrd.open_workbook(str(src_path), formatting_info=False)
    main = rb.sheet_by_name("내역서")
    tok_rows = row_indices_tok(main.nrows)
    jo_rows = row_indices_jo(main.nrows)

    # 내역서 — 필터 복사
    wb_tok = xlwt.Workbook(encoding="utf-8")
    ws_tok = wb_tok.add_sheet("내역서")
    for out_r, src_r in enumerate(tok_rows):
        for c in range(main.ncols):
            copy_cell(main, src_r, c, ws_tok, out_r)
    for sh_name in rb.sheet_names():
        if sh_name == "내역서":
            continue
        sh = rb.sheet_by_name(sh_name)
        ws = wb_tok.add_sheet(sh_name)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                copy_cell(sh, r, c, ws, r)
    wb_tok.save(str(tok_path))

    wb_jo = xlwt.Workbook(encoding="utf-8")
    ws_jo = wb_jo.add_sheet("내역서")
    for out_r, src_r in enumerate(jo_rows):
        for c in range(main.ncols):
            copy_cell(main, src_r, c, ws_jo, out_r)
    for sh_name in rb.sheet_names():
        if sh_name == "내역서":
            continue
        sh = rb.sheet_by_name(sh_name)
        ws = wb_jo.add_sheet(sh_name)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                copy_cell(sh, r, c, ws, r)
    wb_jo.save(str(jo_path))

    return len(tok_rows), len(jo_rows)


def verify_counts() -> None:
    sys.path.insert(0, str(ROOT / "tools"))
    import apply_standard_prices as asp  # noqa: E402

    orig_items, _ = asp.load_estimate(SRC)
    tok_items, _ = asp.load_estimate(OUT_TOK)
    jo_items, _ = asp.load_estimate(OUT_JO)
    print(f"원본 수량행: {len(orig_items)}")
    print(f"토목 수량행: {len(tok_items)} (기대 493)")
    print(f"조경 수량행: {len(jo_items)} (기대 83)")
    if len(tok_items) + len(jo_items) != len(orig_items):
        raise SystemExit("분리 후 합계 건수 불일치")
    if len(tok_items) != 493 or len(jo_items) != 83:
        raise SystemExit("분리 건수 기대값 불일치 — 분리 기준 재확인 필요")


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"원본 없음: {SRC}")
    tok_n, jo_n = split_workbook(SRC, OUT_TOK, OUT_JO)
    print(f"저장: {OUT_TOK.name} ({tok_n}행)")
    print(f"저장: {OUT_JO.name} ({jo_n}행)")
    verify_counts()
    print("검증 OK")


if __name__ == "__main__":
    main()
