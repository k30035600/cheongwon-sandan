#!/usr/bin/env python3
"""포털 탭 ↔ 좌/우 패널 매핑·파일 존재 검사."""
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
html = (ROOT / "청원지구_포털.html").read_text(encoding="utf-8")

# TABS 정의에서 main / placeholder / refs 경로 추출(근사)
paths: set[str] = set()
for m in re.finditer(r'"(0\d_[^"]+\.(?:html|md|pdf|jpg|txt))"', html):
    paths.add(m.group(1))
for m in re.finditer(r'B \+ "/([^"]+)"', html):
    paths.add("05_내역서/" + m.group(1))

missing = [p for p in sorted(paths) if not (ROOT / p).exists()]
print(f"참조 경로 {len(paths)}건 · 누락 {len(missing)}건")
for p in missing:
    print(f"  ✗ {p}")

# 설계 규칙 요약(수동 기대값)
rules = """
탭 그룹별 좌/우 역할(수정 후):
  진명·토지·환평·지구·인허가: 좌 html | 우 pdf/img/txt
  공내역서 목록:           좌 html 안내 | 우 XLS 목록
  공내역서 개별:           좌 안내 placeholder | 우 해당 XLS
  내역서:                  좌 _요약.md | 우 _표준단가산출.xlsx
  집계표 HTML:             좌 총괄 html | 우 xlsx·md
  집계표 xlsx/md:          좌 총괄표.md | 우 xlsx
  미매칭·검토:             좌 안내 placeholder | 우 xlsx
"""
print(rules)
