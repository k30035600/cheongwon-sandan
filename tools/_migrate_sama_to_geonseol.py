#!/usr/bin/env python3
"""삼아산업(개발) → 삼아건설 일괄 치환(텍스트·파일명·xlsx/pptx)."""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]

TEXT_REPLACEMENTS = [
    ("삼아산업개발주식회사", "삼아건설주식회사"),
    ("삼아산업개발", "삼아건설"),
    ("삼아산업", "삼아건설"),
    ("삼아개발", "삼아건설"),
]

TEXT_SUFFIXES = {".py", ".md", ".html", ".js", ".json", ".txt", ".csv"}
SKIP_DIRS = {".git", "__pycache__", "gpu-lab", "node_modules"}
BINARY_SUFFIXES = {".xlsx", ".xlsm", ".pptx", ".docx"}

RENAMES = [
    (ROOT / "07_타견적" / "삼아산업(청원지구).xlsx", ROOT / "07_타견적" / "삼아건설(청원지구).xlsx"),
    (ROOT / "07_타견적" / "삼아산업(청원지구).pdf", ROOT / "07_타견적" / "삼아건설(청원지구).pdf"),
    (ROOT / "07_타견적" / "견적비교(세흥_삼아산업).xlsx", ROOT / "07_타견적" / "견적비교(세흥_삼아건설).xlsx"),
    (
        ROOT / "09_공사지명원" / "지명원_삼아산업개발",
        ROOT / "09_공사지명원" / "지명원_삼아건설",
    ),
    (
        ROOT / "09_공사지명원" / "지명원_삼아건설" / "삼아산업개발 공사지명원.pptx",
        ROOT / "09_공사지명원" / "지명원_삼아건설" / "삼아건설 공사지명원.pptx",
    ),
]


def _apply_text(s: str) -> str:
    for old, new in TEXT_REPLACEMENTS:
        s = s.replace(old, new)
    return s


def patch_text_file(path: Path) -> bool:
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return False
    new_text = _apply_text(text)
    if new_text == text:
        return False
    path.write_text(new_text, encoding="utf-8")
    print(f"  text {path.relative_to(ROOT)}")
    return True


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
        print(f"  zip  {path.relative_to(ROOT)}")
    else:
        tmp.unlink(missing_ok=True)
    return changed


def main() -> int:
    print("=== 삼아산업(개발) → 삼아건설 ===")

    old_dir = ROOT / "09_공사지명원" / "지명원_삼아산업개발"
    new_dir = ROOT / "09_공사지명원" / "지명원_삼아건설"
    if old_dir.is_dir() and not new_dir.exists():
        old_dir.rename(new_dir)
        print(f"  renamed dir {old_dir.name} → {new_dir.name}")

    for old, new in RENAMES:
        if old == old_dir:
            continue
        if old.is_file() and not new.exists():
            new.parent.mkdir(parents=True, exist_ok=True)
            old.rename(new)
            print(f"  renamed {old.relative_to(ROOT)} → {new.name}")

    n_text = 0
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.name == Path(__file__).name:
            continue
        if p.suffix.lower() in TEXT_SUFFIXES:
            if patch_text_file(p):
                n_text += 1
        elif p.suffix.lower() in BINARY_SUFFIXES:
            patch_zip_text(p)

    leftovers = []
    for p in ROOT.rglob("*"):
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if re.search(r"삼아산업|삼아개발", p.name):
            leftovers.append(p.relative_to(ROOT))
    if leftovers:
        print("[잔여 경로명]")
        for x in leftovers:
            print(f"  {x}")
    else:
        print("경로명 잔여 없음")

    # 내용 잔여 검사(텍스트 파일)
    remain = []
    for p in ROOT.rglob("*"):
        if not p.is_file() or p.suffix.lower() not in TEXT_SUFFIXES:
            continue
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.name == Path(__file__).name:
            continue
        try:
            t = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        if re.search(r"삼아산업|삼아개발", t):
            remain.append(p.relative_to(ROOT))
    if remain:
        print("[텍스트 잔여]")
        for x in remain[:20]:
            print(f"  {x}")
    else:
        print(f"텍스트 치환 {n_text}건 · 삼아산업/삼아개발 잔여 0")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
