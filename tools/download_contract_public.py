#!/usr/bin/env python3
"""Download public contract-review files."""
from __future__ import annotations

import sys
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "09_공사지명원" / "외부수집" / "공개자료"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def download(url: str, dest: Path) -> None:
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = resp.read()
        cd = resp.headers.get("Content-Disposition", "")
    dest.write_bytes(data)
    print(f"OK {dest.name} {len(data)} bytes  CD={cd[:120] if cd else '-'}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    items = [
        (
            "https://www.law.go.kr/flDownload.do?"
            + "bylClsCd=200201&flNm=%5B%EB%B3%84%ED%91%9C%5D+%EB%AF%BC%EA%B0%84%EA%B1%B4%EC%84%A4%EA%B3%B5%EC%82%AC+"
            + "%ED%91%9C%EC%A4%80%EB%8F%84%EA%B8%89%EA%B3%84%EC%95%BD%EC%84%9C%28%EC%A0%9C2%EC%A1%B0+%EA%B4%80%EB%A0%A8%29&flSeq=151312547",
            OUT / "01_민간건설공사_표준도급계약서_별표.pdf",
        ),
        (
            "https://law.go.kr/LSW/flDownload.do?"
            + "bylClsCd=200208&flNm=%5B%EB%B6%99%EC%9E%84%5D+%EB%AF%BC%EA%B0%84%EA%B1%B4%EC%84%A4%EA%B3%B5%EC%82%AC+"
            + "%ED%91%9C%EC%A4%80%EB%8F%84%EA%B8%89%EA%B3%84%EC%95%BD%EC%84%9C&flSeq=132844007",
            OUT / "02_민간건설공사_표준도급계약서_붙임.pdf",
        ),
        (
            "https://law.kofia.or.kr/download.do?gubun=101&seq=1723",
            OUT / "06_금융투자협회_책임준공확약_PF대출_모범규준_2025.hwp",
        ),
        (
            "https://law.kofia.or.kr/download.do?gubun=101&seq=1673",
            OUT / "07_금융투자협회_책임준공확약_토지신탁_모범규준_2025.hwp",
        ),
        (
            "https://www.fsc.go.kr/comm/getFile?srvcId=BBSTY1&upperNo=83403&fileTy=ATTACH&fileNo=4",
            OUT / "08_부동산PF_제도개선방안_별첨_2024.pdf",
        ),
        (
            "https://www.fsc.go.kr/comm/getFile?srvcId=BBSTY1&upperNo=83403&fileTy=ATTACH&fileNo=1",
            OUT / "09_부동산PF_제도개선방안_보도자료_2024.pdf",
        ),
    ]
    for url, dest in items:
        try:
            download(url, dest)
        except Exception as e:
            print(f"FAIL {dest.name}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
