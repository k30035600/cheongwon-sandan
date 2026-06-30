#!/usr/bin/env python3
"""최종 산출물(총괄·검토·미매칭·보고) → 05_내역서 루트 이동."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path("05_내역서")
WORK = BASE / "내역서작업"

# 내역서작업 → 05_내역서 루트 (표준단가산출 xlsx/md는 내역서작업에 유지)
TO_ROOT = [
    "총괄표.md",
    "총괄표.xlsx",
    "총괄표_공종별.xlsx",
    "내역서_표준단가산출_총괄표.html",
    "청원지구_종합보고서.html",
    "토공_물량비교표.md",
    "토공_물량비교표.html",
    "검토_전체.xlsx",
    "검토_일위대가산출.xlsx",
    "검토_공종별_일위대가산출.xlsx",
    "검토_토공_일위대가산출.xlsx",
    "검토_품목검증.xlsx",
    "미매칭_전체.xlsx",
    "미매칭_일위대가산출.xlsx",
    "조경수_미매칭점검.xlsx",
    "조달청_미매칭점검.xlsx",
    "조달청보정_매칭결과.xlsx",
    "한국물가협회_미매칭점검.xlsx",
    "토목_미매칭_수동단가입력표.xlsx",
]
TO_ROOT_DIR = [
    ("조경시설물", "조경시설물"),
]


def main() -> None:
    moved = 0
    for name in TO_ROOT:
        src, dst = WORK / name, BASE / name
        if not src.exists():
            if dst.exists():
                continue
            print(f"  skip missing: {name}")
            continue
        if dst.exists() and src.resolve() != dst.resolve():
            dst.unlink()
        shutil.move(str(src), str(dst))
        print(f"  → {name}")
        moved += 1
    for sub, dest in TO_ROOT_DIR:
        src, dst = WORK / sub, BASE / dest
        if not src.exists():
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists():
            shutil.rmtree(dst)
        shutil.move(str(src), str(dst))
        print(f"  → {dest}/")
        moved += 1
    print(f"\n이동 {moved}건 · 내역서작업 잔류: {sum(1 for _ in WORK.rglob('*') if _.is_file())} files")


if __name__ == "__main__":
    main()
