#!/usr/bin/env python3
"""05_내역서 → 공내역서 / 내역서작업 / 일위대가DB / 표준품셈 4분류 정리."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "05_내역서"
GONG = BASE / "공내역서"
WORK = BASE / "내역서작업"
JO_GONG = GONG / "조경시설물"
JO_WORK = WORK / "조경시설물"

MOVES: list[tuple[str, str]] = [
    ("공내역서_안내.html", "공내역서/공내역서_안내.html"),
    ("조경시설물/01_260620화성시마도면청원리청원지구 조경.XLS", "공내역서/조경시설물/01_260620화성시마도면청원리청원지구 조경.XLS"),
    ("조경시설물/README.md", "내역서작업/조경/조경시설물/README.md"),
    ("조경시설물/조경시설물_미매칭_20건.xlsx", "내역서작업/조경/조경시설물/조경시설물_미매칭_20건.xlsx"),
    ("조경시설물/조경시설물_미매칭_20건.csv", "내역서작업/조경/조경시설물/조경시설물_미매칭_20건.csv"),
    ("검토_전체.xlsx", "내역서작업/_공통/검토_전체.xlsx"),
    ("검토_일위대가산출.xlsx", "내역서작업/_공통/검토_일위대가산출.xlsx"),
    ("검토_공종별_일위대가산출.xlsx", "내역서작업/_공통/검토_공종별_일위대가산출.xlsx"),
    ("검토_토공_일위대가산출.xlsx", "내역서작업/토목/검토_토공_일위대가산출.xlsx"),
    ("검토_품목검증.xlsx", "내역서작업/_공통/검토_품목검증.xlsx"),
    ("미매칭_전체.xlsx", "내역서작업/_공통/미매칭_전체.xlsx"),
    ("미매칭_일위대가산출.xlsx", "내역서작업/_공통/미매칭_일위대가산출.xlsx"),
    ("조경수_미매칭점검.xlsx", "내역서작업/조경/조경수_미매칭점검.xlsx"),
    ("조달청_미매칭점검.xlsx", "내역서작업/_공통/조달청_미매칭점검.xlsx"),
    ("조달청보정_매칭결과.xlsx", "내역서작업/_공통/조달청보정_매칭결과.xlsx"),
    ("한국물가협회_미매칭점검.xlsx", "내역서작업/_공통/한국물가협회_미매칭점검.xlsx"),
    ("토목_미매칭_수동단가입력표.xlsx", "내역서작업/토목/토목_미매칭_수동단가입력표.xlsx"),
    ("총괄표.md", "내역서작업/_공통/총괄표.md"),
    ("총괄표.xlsx", "내역서작업/_공통/총괄표.xlsx"),
    ("총괄표_공종별.xlsx", "내역서작업/_공통/총괄표_공종별.xlsx"),
    ("내역서_표준단가산출_총괄표.html", "내역서작업/_공통/내역서_표준단가산출_총괄표.html"),
    ("청원지구_종합보고서.html", "내역서작업/_공통/청원지구_종합보고서.html"),
    ("토공_물량비교표.md", "내역서작업/토목/토공_물량비교표.md"),
    ("토공_물량비교표.html", "내역서작업/토목/토공_물량비교표.html"),
]


def move_one(src_rel: str, dst_rel: str) -> None:
    src, dst = BASE / src_rel, BASE / dst_rel
    if not src.exists():
        if dst.exists():
            print(f"  skip (already): {dst_rel}")
            return
        print(f"  missing: {src_rel}")
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists() and src.resolve() != dst.resolve():
        dst.unlink()
    shutil.move(str(src), str(dst))
    print(f"  {src_rel} → {dst_rel}")


def main() -> None:
    print("=== 05_내역서 4분류 정리 ===")
    for s, d in MOVES:
        move_one(s, d)
    # 빈 조경시설물 루트 폴더 제거
    old_jo = BASE / "공내역서" / "조경시설물"
    if old_jo.exists() and old_jo.is_dir() and not any(old_jo.iterdir()):
        old_jo.rmdir()
        print("  removed empty: 조경시설물/")
    # 루트 잔여 파일 점검
    left = [p.name for p in BASE.iterdir() if p.is_file()]
    if left:
        print(f"\n[주의] 루트 잔여 파일 {len(left)}개: {left}")
    else:
        print("\n루트 산출물 0건 — 4분류만 유지")
    for d in ("공내역서", "내역서작업", "일위대가DB", "표준품셈"):
        n = sum(1 for _ in (BASE / d).rglob("*") if _.is_file())
        print(f"  {d}: {n} files")


if __name__ == "__main__":
    main()
