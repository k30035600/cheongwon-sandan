#!/usr/bin/env python3
"""HNK → 삼아건설 일괄 치환(파일명·xlsx/pptx 내 문자열·비교 스크립트)."""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]

TEXT_REPLACEMENTS = [
    ("HNK산업개발", "삼아건설"),
    ("에이치앤케이산업개발", "삼아건설"),
    ("에이치앤케이", "삼아건설"),
    ("HNK", "삼아건설"),
]

RENAMES = [
    (ROOT / "07_타견적" / "HNK(청원지구).xlsx", ROOT / "07_타견적" / "삼아건설(청원지구).xlsx"),
    (ROOT / "07_타견적" / "HNK(청원지구).pdf", ROOT / "07_타견적" / "삼아건설(청원지구).pdf"),
    (ROOT / "07_타견적" / "견적비교(세흥_HNK).xlsx", ROOT / "07_타견적" / "견적비교(세흥_삼아건설).xlsx"),
    (
        ROOT / "09_공사지명원" / "지명원_HNK산업개발",
        ROOT / "09_공사지명원" / "지명원_삼아건설",
    ),
    (
        ROOT / "09_공사지명원" / "지명원_삼아건설" / "HNK산업개발 공사지명원(가영현).pptx",
        ROOT / "09_공사지명원" / "지명원_삼아건설" / "삼아건설 공사지명원.pptx",
    ),
    (
        ROOT / "tools" / "build_estimate_compare_hnk.py",
        ROOT / "tools" / "build_estimate_compare_sama.py",
    ),
]

CODE_REPLACEMENTS = [
    ("HNK산업개발", "삼아건설"),
    ("에이치앤케이산업개발", "삼아건설"),
    ("세흥_HNK", "세흥_삼아건설"),
    ("세흥−HNK", "세흥−삼아건설"),
    ("HNK(청원지구)", "삼아건설(청원지구)"),
    ("FILL_HNK", "FILL_SAMA"),
    ("read_hnk", "read_sama"),
    ("load_hnk", "load_sama"),
    ("match_hnk", "match_sama"),
    ("hnk_cost_cascade", "sama_cost_cascade"),
    ("hnk_cost", "sama_cost"),
    ("hnk_rates", "sama_rates"),
    ("hnk_direct", "sama_direct"),
    ("hnk_supply", "sama_supply"),
    ("hnk_indirect", "sama_indirect"),
    ("hnk_vat", "sama_vat"),
    ("hnk_dogup", "sama_dogup"),
    ("hnk_rate", "sama_rate"),
    ("hnk_gong", "sama_gong"),
    ("hnk_items", "sama_items"),
    ("hnkv", "samav"),
    ("HNK =", "SAMA ="),
    ("HNK,", "SAMA,"),
    ('"hnk"', '"sama"'),
    ("'hnk'", "'sama'"),
    ("grand[\"hnk\"]", "grand[\"sama\"]"),
    ("sub[\"hnk\"]", "sub[\"sama\"]"),
    ("cls_sub[\"hnk\"]", "cls_sub[\"sama\"]"),
    ("gong_sub[\"hnk\"]", "gong_sub[\"sama\"]"),
    ("build_estimate_compare_hnk", "build_estimate_compare_sama"),
    ("HNK ", "삼아건설 "),
    ("HNK:", "삼아건설:"),
    ("HNK·", "삼아건설·"),
    ("·HNK", "·삼아건설"),
    ("HNK)", "삼아건설)"),
    ("(HNK", "(삼아건설"),
    ("↔ HNK", "↔ 삼아건설"),
    ("−HNK", "−삼아건설"),
    ("금액차이 = 세흥 − HNK", "금액차이 = 세흥 − 삼아건설"),
]


def _apply_text(s: str) -> str:
    for old, new in TEXT_REPLACEMENTS:
        s = s.replace(old, new)
    return s


def patch_zip_text(path: Path) -> bool:
    if not path.is_file():
        return False
    tmp = path.with_suffix(path.suffix + ".tmp")
    changed = False
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    zout.writestr(item, data)
                    continue
                new_text = _apply_text(text)
                if new_text != text:
                    changed = True
                    data = new_text.encode("utf-8")
            zout.writestr(item, data)
    if changed:
        tmp.replace(path)
        print(f"  patched {path.relative_to(ROOT)}")
    else:
        tmp.unlink(missing_ok=True)
    return changed


def migrate_compare_script() -> None:
    dst = ROOT / "tools" / "build_estimate_compare_sama.py"
    if not dst.is_file():
        print("  skip compare script (missing)")
        return
    text = dst.read_text(encoding="utf-8")
    for old, new in CODE_REPLACEMENTS:
        text = text.replace(old, new)
    dst.write_text(text, encoding="utf-8")
    print(f"  patched {dst.relative_to(ROOT)}")


def main() -> int:
    print("=== HNK → 삼아건설 마이그레이션 ===")

    # 1) 디렉터리 rename (파일 rename 전)
    old_dir = ROOT / "09_공사지명원" / "지명원_HNK산업개발"
    new_dir = ROOT / "09_공사지명원" / "지명원_삼아건설"
    if old_dir.is_dir() and not new_dir.exists():
        old_dir.rename(new_dir)
        print(f"  renamed dir {old_dir.name} → {new_dir.name}")

    # 2) 파일 rename
    for old, new in RENAMES:
        if old == new_dir or old == old_dir:
            continue
        if old.is_file() and not new.exists():
            new.parent.mkdir(parents=True, exist_ok=True)
            old.rename(new)
            print(f"  renamed {old.relative_to(ROOT)} → {new.name}")

    # 3) compare script
    migrate_compare_script()

    # 4) xlsx/pptx 문자열
    targets = [
        ROOT / "07_타견적" / "삼아건설(청원지구).xlsx",
        ROOT / "07_타견적" / "견적비교(세흥_삼아건설).xlsx",
        ROOT / "09_공사지명원" / "지명원_삼아건설" / "삼아건설 공사지명원.pptx",
    ]
    for p in targets:
        patch_zip_text(p)

    # 5) 잔여 HNK 파일명 검사
    leftovers = []
    for p in ROOT.rglob("*"):
        if any(x in p.parts for x in (".git", "__pycache__", "gpu-lab")):
            continue
        if re.search(r"HNK|hnk", p.name):
            leftovers.append(p.relative_to(ROOT))
    if leftovers:
        print("  [잔여 HNK 경로]")
        for x in leftovers:
            print(f"    {x}")
    else:
        print("  HNK 파일명 잔여 없음")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
