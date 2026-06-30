#!/usr/bin/env python3
"""일회성 tools 스크립트·임시 파일 정리 → tools/_archive/ · _land_ocr 삭제."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
ARCHIVE = TOOLS / "_archive"

TO_ARCHIVE = [
    # debug / audit
    "_audit_deploy_assets.py",
    "_audit_portal_full.py",
    "_audit_portal_tabs.py",
    "_audit_rowpatch.py",
    "_check_row_align.py",
    "_diag_rowsrc.py",
    "_dump_shopmall.py",
    "_factcheck_confirmed.py",
    "_list_tree_unmatched.py",
    "_probe_ilwidae.py",
    "_probe_shopmall2.py",
    "_scan_lowscore_conf.py",
    "_scan_outliers.py",
    "_scan_rebar.py",
    "_trace_item.py",
    "_trace_row666.py",
    "_validity_check.py",
    "_verify_confirmed.py",
    # folder migration (applied)
    "reorganize_05_naeyeokseo.py",
    "move_final_outputs_to_root.py",
    "restore_05_four_folders.py",
    "patch_05_work_paths.py",
    "patch_05_root_outputs.py",
    # one-off patches / builders
    "patch_qc_months.py",
    "patch_jo_facility_formulas.py",
    "split_01_tok_jo.py",
    "clean_xls_sheets.py",
    "extract_land_register.py",
    "build_pe_tarp_ilwidae.py",
    "build_crane_ilwidae.py",
]

TO_DELETE_DIRS = [
    TOOLS / "_land_ocr",
]

TO_DELETE_FILES = [
    TOOLS / "_poomsem_cache" / "download_xlsx.xlsx",
]


def main() -> None:
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    moved = skipped = 0
    for name in TO_ARCHIVE:
        src = TOOLS / name
        dst = ARCHIVE / name
        if not src.is_file():
            skipped += 1
            continue
        if dst.exists():
            dst.unlink()
        shutil.move(str(src), str(dst))
        print(f"  archive: {name}")
        moved += 1
    for d in TO_DELETE_DIRS:
        if d.is_dir():
            shutil.rmtree(d)
            print(f"  deleted dir: {d.relative_to(ROOT)}")
    for f in TO_DELETE_FILES:
        if f.is_file():
            f.unlink()
            print(f"  deleted: {f.relative_to(ROOT)}")
    print(f"\n완료 — archive {moved}건 · 없음 {skipped}건")


if __name__ == "__main__":
    main()
