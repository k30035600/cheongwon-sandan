#!/usr/bin/env python3
"""산출물 루트 복귀 후 경로 참조 되돌리기 (내역서작업/_공통/총괄표 → 루트)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
WORK = "내역서작업"

# 루트 산출물 — BASE / "내역서작업" / "파일" → BASE / "파일"
ROOT_FILES = [
    "총괄표", "검토_", "미매칭_", "조경수_", "조달청", "한국물가",
    "토목_미매칭", "토공_물량", "내역서_표준", "청원지구_종합",
]

def patch_text(text: str) -> str:
    t = text
    # 조경시설물 산출물은 루트
    t = t.replace(f'05_내역서/{WORK}/조경시설물/', '05_내역서/조경시설물/')
    t = t.replace(f'B + "/{WORK}/조경시설물/', 'B + "/조경시설물/')
    for prefix in ROOT_FILES:
        t = t.replace(f'BASE / "{WORK}" / "{prefix}', f'BASE / "{prefix}')
        t = t.replace(f'05_내역서/{WORK}/{prefix}', f'05_내역서/{prefix}')
        t = t.replace(f'B + "/{WORK}/{prefix}', f'B + "/{prefix}')
    # UM/RV forEach
    t = t.replace(f'B + "/{WORK}/" + file', 'B + "/" + file')
    # build_consolidated OUT_DIR
    t = t.replace(f'OUT_DIR = BASE / "{WORK}"', 'OUT_DIR = BASE')
    t = t.replace(f'OUT_DIR = BASE / "{WORK}" / "{WORK}"', 'OUT_DIR = BASE')
    # export_unmatched
    if 'OUT_DIR = BASE' in t and 'export_unmatched' in str():
        pass
    t = t.replace(f'Path("05_내역서/{WORK}/조경시설물/', 'Path("05_내역서/내역서작업/조경/조경시설물/')
    # 공내역서_안내 stays in 공내역서/
    return t

def patch_file(path: Path) -> bool:
    if not path.exists():
        return False
    old = path.read_text(encoding="utf-8")
    new = patch_text(old)
    if new != old:
        path.write_text(new, encoding="utf-8")
        return True
    return False

def main():
    n = 0
    for pat in ["tools/*.py", "청원지구_포털.html", "08_제출내역서/*.html",
                "대화시작하기.md", "05_내역서/README.md", "05_내역서/총괄표.md",
                "05_내역서/조경시설물/README.md"]:
        if "*" in pat:
            for p in ROOT.glob(pat):
                n += patch_file(p)
        else:
            n += patch_file(ROOT / pat)
    print(f"patched {n} files")

if __name__ == "__main__":
    main()
