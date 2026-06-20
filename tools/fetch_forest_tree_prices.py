#!/usr/bin/env python3
"""forestinfo.or.kr 조경수 관측시세 → CSV 정규화.

공식 접근: 산림조합중앙회·임산물유통정보시스템(forestinfo.or.kr)
  → 가격정보 → 나무 가격정보 → 조경수 관측시세
URL: https://www.forestinfo.or.kr/public/tree_prinfo/selectTreePrinfoFoppList.do

성격: 한국농촌경제원 조경수 임업관측 참고가(고시 아님). 단위 원/주·상차도.
출력: 05_내역서/일위대가DB/조경수_관측시세.csv
"""
from __future__ import annotations

import csv
import re
import ssl
import sys
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "05_내역서" / "일위대가DB" / "조경수_관측시세.csv"
UPLOAD_MD = Path(__file__).resolve().parent.parent / "05_내역서" / "일위대가DB" / "_외부원본" / "forestinfo_조경수관측시세.md"
FALLBACK_MD = Path(r"C:\Users\k3003\.cursor\projects\d-OneDrive-Cursor\uploads\selectTreePrinfoFoppList.do-0.md")
URL = "https://www.forestinfo.or.kr/public/tree_prinfo/selectTreePrinfoFoppList.do"
SOURCE_LABEL = "forestinfo 조경수관측시세(참고가)"
SOURCE_NOTE = (
    "산림조합중앙회·임산물유통정보시스템 → 가격정보 → 나무 가격정보 → 조경수 관측시세"
)
# 표 헤더: 2018~2022 × (하한·상한) = 10열
YEARS = (2018, 2019, 2020, 2021, 2022)

SPEC_CELL_RE = re.compile(
    r"^(?:[HRBW][\d.(]|H[\d.]+\(|R\d|B\d|[\d.]+\)|Ⅹ|X)",
    re.I,
)

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE


def norm_spec(s: str) -> str:
    s = s.strip()
    s = s.replace("×", "X").replace("Ⅹ", "X").replace("x", "X")
    s = re.sub(r"\s+", "", s)
    return s


def is_spec_cell(s: str) -> bool:
    s = (s or "").strip()
    if not s or s == "규격":
        return False
    if SPEC_CELL_RE.match(s):
        return True
    return bool(re.search(r"[ⅩX×]", s))


def append_price_rows(
    rows: list[dict],
    *,
    species: str,
    spec: str,
    cells: list[str],
    price_start: int,
) -> None:
    if not species or not spec:
        return
    for yi, year in enumerate(YEARS):
        lo_i = price_start + yi * 2
        hi_i = lo_i + 1
        if hi_i >= len(cells):
            break
        lo, hi, mid = parse_price(cells[lo_i])
        if mid is None and lo is None and hi is None:
            continue
        if lo is None and hi is None and mid is not None:
            price_lo = price_hi = price_mid = mid
        else:
            price_lo = lo if lo is not None else mid
            price_hi = hi if hi is not None else mid
            if price_lo is None and price_hi is not None:
                price_lo = price_hi
            if price_hi is None and price_lo is not None:
                price_hi = price_lo
            price_mid = (
                (price_lo + price_hi) // 2
                if price_lo and price_hi
                else (mid or price_lo or price_hi)
            )
        rows.append({
            "species": species,
            "spec_raw": spec,
            "spec_norm": norm_spec(spec),
            "year": year,
            "price_lo": price_lo,
            "price_hi": price_hi,
            "price_mid": price_mid,
            "unit": "주",
            "source": SOURCE_LABEL,
        })


def parse_price(tok: str) -> tuple[int | None, int | None, int | None]:
    tok = (tok or "").strip().replace(",", "").replace("\\-", "-")
    if not tok or tok == "-":
        return None, None, None
    m = re.match(r"^\((\d+)\)$", tok)
    if m:
        mid = int(m.group(1))
        return mid, mid, mid
    if tok.isdigit():
        v = int(tok)
        return v, v, v
    return None, None, None


def parse_markdown_table(text: str) -> list[dict]:
    rows: list[dict] = []
    current_species = ""
    in_table = False
    header_done = False
    for line in text.splitlines():
        if "| 품목" in line and "규격" in line:
            in_table = True
            continue
        if not in_table or not line.strip().startswith("|"):
            continue
        if line.strip().startswith("| ---") or "__" in line:
            continue
        cells = [c.strip().replace("\\-", "-") for c in line.strip().strip("|").split("|")]
        if not header_done:
            if cells and cells[0] in ("2018", "2019", "2020", "2021", "2022", "하한가(중간가격)"):
                continue
            if cells and cells[0] == "품목":
                continue
            header_done = True
        if len(cells) < 12:
            continue
        c0, c1 = cells[0], cells[1]
        # 연속 행: col0=규격(H/R/B…), col1부터 가격
        if is_spec_cell(c0):
            if not current_species:
                continue
            append_price_rows(
                rows,
                species=current_species,
                spec=c0,
                cells=cells,
                price_start=1,
            )
            continue
        if not c0:
            if is_spec_cell(c1) and current_species:
                append_price_rows(
                    rows,
                    species=current_species,
                    spec=c1,
                    cells=cells,
                    price_start=2,
                )
            continue
        current_species = c0
        if not is_spec_cell(c1):
            continue
        append_price_rows(
            rows,
            species=current_species,
            spec=c1,
            cells=cells,
            price_start=2,
        )
    return rows


def parse_html_table(html: str) -> list[dict]:
    """HTML table fallback — markdown 우선."""
    rows: list[dict] = []
    # species rows: <td>느티나무</td><td>R12</td>...
    trs = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I)
    current_species = ""
    for tr in trs:
        tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)
        if len(tds) < 12:
            continue
        def clean(x):
            return re.sub(r"<[^>]+>", "", x).strip().replace(",", "")
        cells = [clean(x) for x in tds]
        if cells[0] == "품목":
            continue
        c0, c1 = cells[0], cells[1]
        if is_spec_cell(c0):
            if not current_species:
                continue
            append_price_rows(
                rows,
                species=current_species,
                spec=c0,
                cells=cells,
                price_start=1,
            )
            continue
        if not c0:
            if is_spec_cell(c1) and current_species:
                append_price_rows(
                    rows,
                    species=current_species,
                    spec=c1,
                    cells=cells,
                    price_start=2,
                )
            continue
        current_species = c0
        if not is_spec_cell(c1):
            continue
        append_price_rows(
            rows,
            species=current_species,
            spec=c1,
            cells=cells,
            price_start=2,
        )
    return rows


def load_source() -> str:
    for p in (UPLOAD_MD, FALLBACK_MD):
        if p.exists():
            return p.read_text(encoding="utf-8")
    req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
        return r.read().decode("utf-8", "replace")


def main():
    text = load_source()
    if "<table" in text.lower() or "<td" in text.lower():
        rows = parse_html_table(text)
    else:
        rows = parse_markdown_table(text)
    if not rows:
        raise SystemExit("표 파싱 실패")
    # 최신 연도 우선 dedupe (species+spec_norm → max year)
    best: dict[tuple[str, str], dict] = {}
    for r in rows:
        key = (r["species"], r["spec_norm"])
        if key not in best or r["year"] > best[key]["year"]:
            best[key] = r
    out_rows = sorted(best.values(), key=lambda x: (x["species"], x["spec_norm"]))
    OUT.parent.mkdir(parents=True, exist_ok=True)
    fields = ["species", "spec_raw", "spec_norm", "year", "price_lo", "price_hi", "price_mid", "unit", "source"]
    with OUT.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(out_rows)
    y2022 = sum(1 for r in out_rows if r["year"] == 2022)
    print(f"출처: {SOURCE_NOTE}")
    print(f"URL: {URL}")
    print(f"저장: {OUT} ({len(out_rows)}건, 2022년 {y2022}건)")


if __name__ == "__main__":
    main()
