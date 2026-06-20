#!/usr/bin/env python3
"""포털 enc() 경로 vs git 추적 — 배포(Railway) 누락 전수조사."""
from __future__ import annotations
import re
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
html = (ROOT / "청원지구_포털.html").read_text(encoding="utf-8")
paths = sorted(set(re.findall(r"""enc\(['"]([^'"]+)['"]\)""", html)))
tracked = set(
    subprocess.check_output(["git", "ls-files"], cwd=ROOT, text=True, encoding="utf-8")
    .splitlines()
)

by_ext: dict[str, list[str]] = {}
for p in paths:
    if p.startswith("http"):
        continue
    if p in tracked:
        continue
    ext = Path(p).suffix.lower() or "(none)"
    by_ext.setdefault(ext, []).append(p)

print(f"포털 enc() 로컬 경로: {sum(1 for p in paths if not p.startswith('http'))}건")
print(f"git 미추적(배포 누락): {sum(len(v) for v in by_ext.values())}건\n")
for ext in sorted(by_ext):
    print(f"=== {ext} ({len(by_ext[ext])}건) ===")
    for p in by_ext[ext]:
        exists = (ROOT / p).is_file()
        print(f"  {'[로컬있음]' if exists else '[로컬없음]'} {p}")
