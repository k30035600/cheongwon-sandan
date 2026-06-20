#!/usr/bin/env python3
"""한국표준품셈정보원 자료실(kseis.co.kr) — 2026년 게시물 첨부 일괄 다운로드.

자료실 목록 HTML을 파싱해 제목에 '2026'이 포함된 게시물의 모든 첨부파일을
05_내역서/한국표준품셈정보원 폴더로 내려받는다(이미 있으면 건너뜀).

  ※ 표준품셈 e-book·표준시장단가 e-book·조달청 가격정보·노임단가는 온라인 열람
    뷰어(유료, 페이지 단위)라 파일 다운로드 대상이 아니다. 실제 파일은 자료실 첨부뿐.
"""
from __future__ import annotations

import re
import sys
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / "05_내역서" / "한국표준품셈정보원"
LIST_URL = "https://www.kseis.co.kr/bbs/data/dataList.do?pgno={pg}"
DOWN = "https://www.kseis.co.kr/bbs/data/dataFileDown.do?bbs_seq={seq}&file_no={fno}"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"

PAGES = (1, 2)          # 2026 게시물은 1~2페이지에 집중
TITLE_RE = re.compile(r'dataDetail\.do\?bbs_seq=(\d+)&(?:amp;)?pgno=\d+">([^<]+)</a>')
FILE_RE = re.compile(r'dataFileDown\.do\?bbs_seq=(\d+)&(?:amp;)?file_no=(\d+)"[^<]*<img[^>]*alt="([^"]+)"')


def get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def sanitize(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def download(seq: str, fno: str, name: str) -> str:
    dest = DEST / sanitize(name)
    if dest.exists() and dest.stat().st_size > 1000:
        return f"[skip] {name}"
    try:
        data = get(DOWN.format(seq=seq, fno=fno))
        if len(data) < 1000 or data[:15].lstrip()[:5].lower() in (b"<!doc", b"<html"):
            return f"[FAIL] {name}: HTML/빈 응답({len(data)}B)"
        dest.write_bytes(data)
        return f"[OK]   {name} ({len(data):,}B)"
    except Exception as e:  # noqa: BLE001
        return f"[FAIL] {name}: {type(e).__name__} {e}"


def main() -> None:
    DEST.mkdir(parents=True, exist_ok=True)
    titles: dict[str, str] = {}
    files: list[tuple[str, str, str]] = []
    for pg in PAGES:
        html = get(LIST_URL.format(pg=pg)).decode("utf-8", "replace")
        for seq, title in TITLE_RE.findall(html):
            titles[seq] = title.strip()
        for seq, fno, name in FILE_RE.findall(html):
            files.append((seq, fno, name))

    todo = [(s, f, n) for s, f, n in files if "2026" in titles.get(s, "")]
    print(f"자료실 2026 게시물 첨부 {len(todo)}개 (제목 기준)\n")
    for seq, fno, name in todo:
        print(download(seq, fno, name))


if __name__ == "__main__":
    main()
