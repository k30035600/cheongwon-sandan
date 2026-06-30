#!/usr/bin/env python3
"""05_내역서 — 4분류만 유지(루트 README 제외, 산출물→내역서작업)."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE = Path("05_내역서")
WORK = BASE / "내역서작업"

TO_WORK = [
    "총괄표.md", "총괄표.xlsx", "총괄표_공종별.xlsx",
    "내역서_표준단가산출_총괄표.html", "청원지구_종합보고서.html",
    "토공_물량비교표.md", "토공_물량비교표.html",
    "검토_전체.xlsx", "검토_일위대가산출.xlsx", "검토_공종별_일위대가산출.xlsx",
    "검토_토공_일위대가산출.xlsx", "검토_품목검증.xlsx",
    "미매칭_전체.xlsx", "미매칭_일위대가산출.xlsx",
    "조경수_미매칭점검.xlsx", "조달청_미매칭점검.xlsx",
    "조달청보정_매칭결과.xlsx", "한국물가협회_미매칭점검.xlsx",
    "토목_미매칭_수동단가입력표.xlsx",
]


def main() -> None:
    WORK.mkdir(parents=True, exist_ok=True)
    n = 0
    for name in TO_WORK:
        src, dst = BASE / name, WORK / name
        if src.is_file():
            if dst.exists():
                dst.unlink()
            shutil.move(str(src), str(dst))
            n += 1
    jo_root = BASE / "조경시설물"
    jo_work = WORK / "조경시설물"
    if jo_root.is_dir():
        jo_work.mkdir(parents=True, exist_ok=True)
        for f in jo_root.iterdir():
            t = jo_work / f.name
            if t.exists():
                t.unlink()
            shutil.move(str(f), str(t))
        try:
            jo_root.rmdir()
        except OSError:
            pass
        n += 1
    root_files = [p.name for p in BASE.iterdir() if p.is_file() and p.name != "README.md"]
    root_dirs = [p.name for p in BASE.iterdir() if p.is_dir()]
    print(f"내역서작업으로 이동 {n}건")
    print(f"루트 파일(README 외): {root_files}")
    print(f"루트 폴더: {root_dirs}")


if __name__ == "__main__":
    main()
