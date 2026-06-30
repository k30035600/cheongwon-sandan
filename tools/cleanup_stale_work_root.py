#!/usr/bin/env python3
"""내역서작업 루트에 남은 공종별 이동 전 중복 파일 삭제."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

WORK = Path("05_내역서/내역서작업")
GONGJONG = {"회전교차로", "진입도로", "토목", "조경", "폐기물", "전기", "_공통"}


def main() -> None:
    if not WORK.is_dir():
        print("없음:", WORK)
        return
    n = 0
    for item in list(WORK.iterdir()):
        if item.name in GONGJONG:
            continue
        if item.is_dir() and item.name == "조경시설물":
            dst = WORK / "조경" / "조경시설물"
            if dst.exists():
                shutil.rmtree(item)
            else:
                shutil.move(str(item), str(dst))
            print("  dir:", item.name)
            n += 1
        elif item.is_file():
            item.unlink()
            print("  file:", item.name)
            n += 1
    print(f"정리 {n}건")


if __name__ == "__main__":
    main()
