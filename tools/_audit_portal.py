#!/usr/bin/env python3
"""Audit portal: verify all referenced local files exist."""
from __future__ import annotations

import re
import sys
from pathlib import Path
from urllib.parse import unquote

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
PORTAL = ROOT / "청원지구_포털.html"


def collect_paths(text: str) -> set[str]:
    paths: set[str] = set()
    # enc("...") and href="..."
    for m in re.findall(r'enc\(\s*"([^"]+)"\s*\)', text):
        paths.add(m)
    for m in re.findall(r'href="([^"]+)"', text):
        paths.add(m)
    for m in re.findall(r'src="([^"]+)"', text):
        paths.add(m)
    return paths


def is_local(p: str) -> bool:
    if p.startswith(("http://", "https://", "#", "mailto:", "javascript:", "${")):
        return False
    if "${" in p:
        return False
    return True


def normalize(p: str) -> str:
    p = p.split("#")[0].split("?")[0]
    return unquote(p)


def main() -> None:
    text = PORTAL.read_text(encoding="utf-8")

    # Gather array-literal refs like [B + "/..", ...] and workPath/kaRef are JS; handle B paths
    # Extract quoted strings that look like file paths with known extensions
    ext_re = re.compile(
        r'"((?:[^"<>]*?)\.(?:html|xlsx|xls|XLS|pdf|md|txt|jpg|jpeg|png|js|hwp|hwpx))"'
    )
    arr_paths = set(ext_re.findall(text))

    # B prefix and workPath
    B = "05_내역서"
    resolved: set[str] = set()
    for p in arr_paths:
        resolved.add(p)
    # B + "/x" patterns
    for m in re.findall(r'B\s*\+\s*"([^"]+)"', text):
        resolved.add(B + m)
    # workPath(gj, file) -> 05_내역서/내역서작업/gj/file
    for gj, file in re.findall(r'workPath\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)', text):
        resolved.add(f"{B}/내역서작업/{gj}/{file}")
    # JMY_BASE + "/x"
    for m in re.findall(r'JMY_BASE\s*\+\s*"([^"]+)"', text):
        resolved.add("09_공사지명원" + m)
    # base + "/공개자료/x" where base = JMY_BASE + "/" + CONTRACT_EXT
    for m in re.findall(r'base\s*\+\s*"([^"]+)"', text):
        resolved.add("09_공사지명원/외부수집" + m)
    # GONG_ITEMS files -> 05_내역서/공내역서/<file>
    for f in re.findall(r'\["[^"]+",\s*"([^"]+\.(?:XLS|xlsx))"', text):
        resolved.add(f"{B}/공내역서/{f}")
    # KA_XLSX
    resolved.add("08_제출내역서/청원지구_단가통합(전기제외)_ka.xlsx")

    # Also direct enc/href/src local paths
    for p in collect_paths(text):
        if is_local(p):
            resolved.add(p)

    missing = []
    ok = []
    for p in sorted(resolved):
        if not is_local(p) or "${" in p or "+" in p:
            continue
        np = normalize(p)
        fp = ROOT / np
        if fp.exists():
            ok.append(np)
        else:
            missing.append(np)

    print(f"# 검사 대상 로컬 경로: {len(ok)+len(missing)}개")
    print(f"# 존재: {len(ok)}  /  누락: {len(missing)}\n")
    if missing:
        print("## ❌ 누락 파일")
        for m in missing:
            print("  -", m)
    else:
        print("## ✅ 누락 없음")
    print("\n## 존재 확인")
    for o in ok:
        print("  -", o)


if __name__ == "__main__":
    main()
