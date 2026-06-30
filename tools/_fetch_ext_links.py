#!/usr/bin/env python3
"""Scrape download links for external collection."""
import re
import sys
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read().decode("utf-8", "replace")


def main() -> None:
    pages = [
        ("kofia-pf", "https://law.kofia.or.kr/service/law/lawFullScreen.do?historySeq=1723&seq=376"),
        ("kofia-trust", "https://law.kofia.or.kr/service/law/lawFullScreen.do?historySeq=1673&seq=364"),
        ("molit-pf", "https://www.molit.go.kr/USR/NEWS/m_71/dtl.jsp?id=95090369"),
        ("fsc-pf", "https://www.fsc.go.kr/no010101/83403"),
    ]
    for name, url in pages:
        print(f"\n=== {name} ===")
        try:
            html = fetch(url)
        except Exception as e:
            print("ERR", e)
            continue
        for pat in [
            r'href="([^"]*(?:download|Download|hwp|HWP|pdf|PDF|file|attach)[^"]*)"',
            r"href='([^']*(?:download|Download|hwp|HWP|pdf|PDF|file|attach)[^']*)'",
            r"(/service/law/[^\"'\s>]+)",
            r"(/USR/[^\"'\s>]+\.(?:pdf|hwp|hwpx))",
            r"(https?://[^\"'\s>]+\.(?:pdf|hwp|hwpx))",
        ]:
            for m in sorted(set(re.findall(pat, html, re.I))):
                if len(m) < 300:
                    print(m)


if __name__ == "__main__":
    main()
