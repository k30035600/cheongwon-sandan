#!/usr/bin/env python3
"""XLS에서 전역변수 시트·빈 시트 제거(내역서 등 나머지 시트 유지)."""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import xlrd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_REMOVE_NAMES = {"전역변수"}


def sheet_has_data(sh: xlrd.sheet.Sheet) -> bool:
    if sh.nrows == 0 or sh.ncols == 0:
        return False
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            if sh.cell_type(r, c) != xlrd.XL_CELL_EMPTY:
                val = sh.cell_value(r, c)
                if val not in ("", None):
                    return True
    return False


def sheets_to_remove(path: Path, extra_names: set[str] | None = None) -> list[str]:
    remove_names = DEFAULT_REMOVE_NAMES | (extra_names or set())
    rb = xlrd.open_workbook(str(path), formatting_info=False)
    removed: list[str] = []
    for name in rb.sheet_names():
        if name in remove_names:
            removed.append(name)
            continue
        sh = rb.sheet_by_name(name)
        if not sheet_has_data(sh):
            removed.append(name)
    return removed


def delete_sheets_com(path: Path, names: list[str]) -> None:
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(path.resolve()))
        for name in names:
            wb.Worksheets(name).Delete()
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()


def clean_file(path: Path, *, backup: bool = True) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(path)
    removed = sheets_to_remove(path)
    if not removed:
        return []
    if backup:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = path.with_suffix(path.suffix + f".bak_{stamp}")
        shutil.copy2(path, bak)
        print(f"  백업: {bak.name}")
    delete_sheets_com(path, removed)
    return removed


def main() -> None:
    parser = argparse.ArgumentParser(description="XLS 전역변수·빈 시트 제거")
    parser.add_argument(
        "files",
        nargs="*",
        type=Path,
        default=[
            ROOT / "05_내역서" / "공내역서" / "01_화성 청원지구 토목.XLS",
            ROOT / "05_내역서" / "공내역서" / "01_화성 청원지구 조경.XLS",
        ],
    )
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()

    for path in args.files:
        path = path if path.is_absolute() else ROOT / path
        print(f"처리: {path.name}")
        removed = clean_file(path, backup=not args.no_backup)
        if removed:
            print(f"  제거: {', '.join(removed)}")
        else:
            print("  제거 대상 없음")
        rb = xlrd.open_workbook(str(path), formatting_info=False)
        print(f"  잔여 시트: {', '.join(rb.sheet_names())}")


if __name__ == "__main__":
    main()
