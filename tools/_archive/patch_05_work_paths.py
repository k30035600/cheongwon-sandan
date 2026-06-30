#!/usr/bin/env python3
"""05_내역서 경로 참조 일괄 갱신 (4분류 정리 후)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
WORK = "내역서작업"
GONG_JO = "공내역서/조경시설물"

# (glob, [(old, new), ...])
RULES: list[tuple[str, list[tuple[str, str]]]] = [
    ("tools/*.py", [
        ('BASE / "내역서작업" / "검토_', f'BASE / "{WORK}" / "검토_'),
        ('BASE / "내역서작업" / "_공통" / "미매칭_', f'BASE / "{WORK}" / "미매칭_'),
        ('BASE / "내역서작업" / "조경수_', f'BASE / "{WORK}" / "조경수_'),
        ('BASE / "내역서작업" / "조달청', f'BASE / "{WORK}" / "조달청'),
        ('BASE / "내역서작업" / "한국물가', f'BASE / "{WORK}" / "한국물가'),
        ('BASE / "내역서작업" / "토목_미매칭', f'BASE / "{WORK}" / "토목_미매칭'),
        ('BASE / "내역서작업" / "총괄표', f'BASE / "{WORK}" / "총괄표'),
        ('OUT_DIR = BASE\n', f'OUT_DIR = BASE / "{WORK}"\n'),
        ('BASE / "공내역서" / "조경시설물"', f'BASE / "공내역서" / "조경시설물"'),
        ('Path("05_내역서/내역서작업/조경/조경시설물/', f'Path("05_내역서/{WORK}/조경시설물/'),
    ]),
    ("tools/build_consolidated_summary.py", [
        ("OUT_DIR = BASE", f'OUT_DIR = BASE / "{WORK}"'),
    ]),
    ("tools/build_summary_html.py", [
        ('SRC = BASE / "내역서작업" / "총괄표', f'SRC = BASE / "{WORK}" / "총괄표'),
        ('OUT = BASE / "내역서_표준', f'OUT = BASE / "{WORK}" / "내역서_표준'),
    ]),
    ("tools/export_unmatched_review.py", [
        ("OUT_DIR = BASE", f'OUT_DIR = BASE / "{WORK}"'),
        ("05_내역서 루트로 통일", f'05_내역서/{WORK}로 통일'),
    ]),
    ("tools/apply_jo_forest_prices.py", [
        ("05_내역서/조경시설물/", f"05_내역서/{GONG_JO}/"),
    ]),
    ("tools/extract_jo_planting.py", [
        ('/ "조경시설물" /', f'/ "공내역서" / "조경시설물" /'),
    ]),
    ("tools/build_dan_ga_ban_yeong.py", [
        ('BASE / "공내역서" / "조경시설물"', f'BASE / "공내역서" / "조경시설물"'),
        ('"조경시설물/', f'"{GONG_JO}/'),
    ]),
    ("청원지구_포털.html", [
        ('05_내역서/공내역서_안내.html', "05_내역서/공내역서/공내역서_안내.html"),
        ('05_내역서/토공_물량비교표', f"05_내역서/{WORK}/토공_물량비교표"),
        ('B + "/_공통/총괄표', f'B + "/{WORK}/총괄표'),
        ('B + "/_공통/내역서_표준', f'B + "/{WORK}/내역서_표준'),
        ('B + "/_공통/청원지구_종합', f'B + "/{WORK}/청원지구_종합'),
        ('B + "/_공통/검토_', f'B + "/{WORK}/검토_'),
        ('B + "/_공통/미매칭_', f'B + "/{WORK}/미매칭_'),
        ('B + "/" + file', f'B + "/{WORK}/" + file'),
    ]),
    ("08_제출내역서/청원지구_공사공정표.html", [
        ("05_내역서/토공_물량비교표", f"05_내역서/{WORK}/토공_물량비교표"),
    ]),
    ("대화시작하기.md", [
        ("조경시설물/", f"{GONG_JO}/"),
        ("05_내역서/총괄표.md", f"05_내역서/{WORK}/총괄표.md"),
    ]),
    (f"05_내역서/{WORK}/조경시설물/README.md", [
        ("05_내역서/조경시설물/", f"05_내역서/{GONG_JO}/"),
        ("조경시설물/조경시설물_미매칭", f"{WORK}/조경시설물/조경시설물_미매칭"),
    ]),
    (f"05_내역서/{WORK}/총괄표.md", [
        (r"루트 `05_내역서\\`에는", f"`05_내역서\\\\{WORK}\\\\`에는"),
    ]),
    ("tools/compare_cost_rates.py", [
        ("05_내역서/총괄표.xlsx", f"05_내역서/{WORK}/총괄표.xlsx"),
    ]),
]

# docstring-only paths in tools
DOC_REPLACEMENTS = [
    ("05_내역서/내역서작업/검토_", f"05_내역서/{WORK}/검토_"),
    ("05_내역서/내역서작업/_공통/미매칭_", f"05_내역서/{WORK}/미매칭_"),
    ("05_내역서/내역서작업/조경수_", f"05_내역서/{WORK}/조경수_"),
    ("05_내역서/내역서작업/조달청", f"05_내역서/{WORK}/조달청"),
    ("05_내역서/내역서작업/_공통/한국물가", f"05_내역서/{WORK}/한국물가"),
]


def patch_file(path: Path, pairs: list[tuple[str, str]]) -> int:
    if not path.exists():
        return 0
    text = path.read_text(encoding="utf-8")
    orig = text
    for old, new in pairs:
        text = text.replace(old, new)
    if text != orig:
        path.write_text(text, encoding="utf-8")
        return 1
    return 0


def main() -> None:
    n = 0
    for pattern, pairs in RULES:
        if "*" in pattern:
            for p in ROOT.glob(pattern):
                n += patch_file(p, pairs)
        else:
            n += patch_file(ROOT / pattern, pairs)
    # docstrings in all tools
    for p in ROOT.glob("tools/*.py"):
        n += patch_file(p, DOC_REPLACEMENTS)
    print(f"patched {n} files")


if __name__ == "__main__":
    main()
