#!/usr/bin/env python3
"""대한건설협회 건설업 임금실태조사(시중노임단가) PDF → 직종별 노임 CSV.

2026년 상반기 적용(2026.1.1~) 개별직종노임단가. 품셈(공량) × 노임 → 노무비 산정 및
표준시장단가 노무비 검증용 참조 테이블. (BOQ 품명 매칭 풀에는 넣지 않는다.)

출력: 05_내역서/일위대가DB/시중노임_2026.csv
컬럼: 번호,직종명,2026.1.1,2025.9.1,2025.1.1,2024.9.1,비고
"""
from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

import fitz

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "05_내역서"
SRC_PDF = BASE / "한국표준품셈정보원" / "2026년_상반기_적용_건설업_임금실태조사_보고서.pdf"
OUT_CSV = BASE / "일위대가DB" / "시중노임_2026.csv"

CODE_RE = re.compile(r"^\d{4}$")
WAGE_RE = re.compile(r"^\d{1,3}(,\d{3})+$")


def parse_wage_table(text: str) -> list[dict]:
    """세로로 분해된 텍스트 토큰을 직종 레코드로 재구성."""
    tokens = [t.strip() for t in text.splitlines() if t.strip()]
    rows: list[dict] = []
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        mark = ""
        m = re.match(r"^(\*{1,2})(\d{4})$", tok)
        if m:
            mark, code = m.group(1), m.group(2)
        elif CODE_RE.match(tok):
            code = tok
        else:
            i += 1
            continue
        # 직종명: 다음 임금(또는 '-')이 나오기 전까지의 한글 토큰
        j = i + 1
        name_chars: list[str] = []
        wages: list[str] = []
        while j < n:
            t = tokens[j]
            if WAGE_RE.match(t):
                wages.append(t.replace(",", ""))
            elif t == "-":
                wages.append("")
            elif CODE_RE.match(t) or re.match(r"^\*{1,2}\d{4}$", t):
                break
            elif len(wages) == 0:
                name_chars.append(t)
            else:
                break
            j += 1
            if len(wages) >= 4:
                break
        name = "".join(name_chars)
        if name and wages:
            rows.append({
                "번호": code,
                "직종명": name,
                "2026.1.1": wages[0] if len(wages) > 0 else "",
                "2025.9.1": wages[1] if len(wages) > 1 else "",
                "2025.1.1": wages[2] if len(wages) > 2 else "",
                "2024.9.1": wages[3] if len(wages) > 3 else "",
                "비고": "조사현장5개미만" if mark == "*" else ("미조사" if mark == "**" else ""),
            })
        i = j
    return rows


def main() -> None:
    doc = fitz.open(SRC_PDF)
    all_rows: list[dict] = []
    for pno in range(doc.page_count):
        t = doc[pno].get_text()
        if "직종명" in t and re.search(r"\d{1,3},\d{3}", t):
            all_rows.extend(parse_wage_table(t))
    doc.close()

    # 번호 중복 제거(첫 등장 유지)
    seen: set[str] = set()
    uniq: list[dict] = []
    for r in all_rows:
        if r["번호"] in seen:
            continue
        seen.add(r["번호"])
        uniq.append(r)

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(
            f, fieldnames=["번호", "직종명", "2026.1.1", "2025.9.1", "2025.1.1", "2024.9.1", "비고"]
        )
        w.writeheader()
        w.writerows(uniq)
    print(f"시중노임 저장: {OUT_CSV} ({len(uniq)}직종)")
    for r in uniq[:6]:
        print(f"  {r['번호']} {r['직종명']}: {r['2026.1.1']}")


if __name__ == "__main__":
    main()
