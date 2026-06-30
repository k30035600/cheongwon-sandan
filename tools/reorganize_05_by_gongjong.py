#!/usr/bin/env python3
"""05_내역서 — 표준품셈·내역서작업·일위대가DB 공종별 하위 분류."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "05_내역서"
WORK = BASE / "내역서작업"
ILW = BASE / "일위대가DB"
POOM = BASE / "표준품셈"
COMMON = "_공통"
GONGJONG = ("회전교차로", "진입도로", "토목", "조경", "폐기물", "전기", COMMON)

# 내역서작업 — 파일명 → 공종
WORK_MAP: dict[str, str] = {
    "05_화성 청원로(회전교차로)_표준단가산출.xlsx": "회전교차로",
    "05_화성 청원로(회전교차로)_표준단가산출_요약.md": "회전교차로",
    "04_화성 청원지구 진입도로 실시설계_표준단가산출.xlsx": "진입도로",
    "04_화성 청원지구 진입도로 실시설계_표준단가산출_요약.md": "진입도로",
    "01_화성 청원지구 토목_표준단가산출.xlsx": "토목",
    "01_화성 청원지구 토목_표준단가산출_요약.md": "토목",
    "06_화성 청원지구 산업유통형 개발행위_표준단가산출.xlsx": "토목",
    "06_화성 청원지구 산업유통형 개발행위_표준단가산출_요약.md": "토목",
    "토공_물량비교표.md": "토목",
    "토공_물량비교표.html": "토목",
    "검토_토공_일위대가산출.xlsx": "토목",
    "검토_토공_품셈산출.xlsx": "토목",
    "토목_미매칭_수동단가입력표.xlsx": "토목",
    "PE천막지_일위대가_10㎡기준.xlsx": "토목",
    "크레인10톤_일위대가_6일기준.xlsx": "토목",
    "임목파쇄_단가기준_체크리스트.md": "토목",
    "01_화성 청원지구 조경_표준단가산출.xlsx": "조경",
    "01_화성 청원지구 조경_표준단가산출_요약.md": "조경",
    "01_260620_푸른조경_식재내역.md": "조경",
    "조경수_미매칭점검.xlsx": "조경",
    "07_화성 청원지구 건설폐기물처리_표준단가산출.xlsx": "폐기물",
    "07_화성 청원지구 건설폐기물처리_표준단가산출_요약.md": "폐기물",
    "02_화성 청원지구 전기설비_표준단가산출.xlsx": "전기",
    "02_화성 청원지구 전기설비_표준단가산출_요약.md": "전기",
    "총괄표.md": COMMON,
    "총괄표.xlsx": COMMON,
    "총괄표_공종별.xlsx": COMMON,
    "내역서_표준단가산출_총괄표.html": COMMON,
    "청원지구_종합보고서.html": COMMON,
    "검토_전체.xlsx": COMMON,
    "검토_일위대가산출.xlsx": COMMON,
    "검토_공종별_일위대가산출.xlsx": COMMON,
    "검토_공종별_품셈산출.xlsx": COMMON,
    "검토_오매칭교정_품셈산출.xlsx": COMMON,
    "검토_품목검증.xlsx": COMMON,
    "미매칭_전체.xlsx": COMMON,
    "미매칭_일위대가산출.xlsx": COMMON,
    "조달청_미매칭점검.xlsx": COMMON,
    "조달청보정_매칭결과.xlsx": COMMON,
    "한국물가협회_미매칭점검.xlsx": COMMON,
}

WORK_DIRS: dict[str, str] = {
    "조경시설물": "조경",
}


def classify_poomsem(name: str) -> str:
    n = name.lower()
    if any(k in name for k in ("전기", "정보통신", "SW기술")):
        return "전기"
    if "조경" in name:
        return "조경"
    if any(
        k in name
        for k in (
            "토목", "토공", "도로포장", "하천", "터널", "관부설", "항만",
            "지반조", "건설기계", "궤도", "기초공사", "강구조", "철근콘",
            "돌공", "시설공통자재(토목)",
        )
    ):
        return "토목"
    return COMMON


def classify_ilw(name: str) -> str:
    if name.startswith("전기(02)_"):
        return "전기"
    if name.startswith("폐기물(07)_"):
        return "폐기물"
    if "조경" in name:
        return "조경"
    return COMMON


def ensure_dirs(base: Path) -> None:
    for g in GONGJONG:
        (base / g).mkdir(parents=True, exist_ok=True)


def move_item(src: Path, dst: Path, dry: bool = False) -> bool:
    if not src.exists():
        return False
    if dst.exists() and src.resolve() == dst.resolve():
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        if dst.is_dir():
            shutil.rmtree(dst)
        else:
            dst.unlink()
    if dry:
        print(f"  [dry] {src.relative_to(ROOT)} → {dst.relative_to(ROOT)}")
    else:
        shutil.move(str(src), str(dst))
        print(f"  {src.relative_to(ROOT)} → {dst.relative_to(ROOT)}")
    return True


def reorganize_work(dry: bool = False) -> int:
    ensure_dirs(WORK)
    n = 0
    # already in gongjong subdir — skip
    for gj in GONGJONG:
        sub = WORK / gj
        if not sub.is_dir():
            continue
    for item in list(WORK.iterdir()):
        if item.name in GONGJONG:
            continue
        if item.is_dir() and item.name in WORK_DIRS:
            dst = WORK / WORK_DIRS[item.name] / item.name
            if move_item(item, dst, dry):
                n += 1
            continue
        if item.is_file() and item.name in WORK_MAP:
            dst = WORK / WORK_MAP[item.name] / item.name
            if move_item(item, dst, dry):
                n += 1
        elif item.is_file():
            print(f"  [skip] 내역서작업/{item.name} (매핑 없음)")
    return n


def reorganize_poomsem(dry: bool = False) -> int:
    ensure_dirs(POOM)
    n = 0
    for item in list(POOM.iterdir()):
        if item.name in GONGJONG:
            continue
        gj = classify_poomsem(item.name)
        dst = POOM / gj / item.name
        if move_item(item, dst, dry):
            n += 1
    return n


def reorganize_ilw(dry: bool = False) -> int:
    ensure_dirs(ILW)
    n = 0
    for item in list(ILW.iterdir()):
        if item.name in GONGJONG or item.name == "README.md":
            continue
        gj = classify_ilw(item.name)
        dst = ILW / gj / item.name
        if move_item(item, dst, dry):
            n += 1
    return n


def main() -> None:
    dry = "--dry" in sys.argv
    print("=== 공종별 하위 분류 ===\n")
    w = reorganize_work(dry)
    p = reorganize_poomsem(dry)
    i = reorganize_ilw(dry)
    print(f"\n이동: 내역서작업 {w} · 표준품셈 {p} · 일위대가DB {i}")


if __name__ == "__main__":
    main()
