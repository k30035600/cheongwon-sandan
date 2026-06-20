"""조달청 나라장터 가격정보현황서비스(15129415) 수집 → 로컬 CSV.

가이드 v1.5(PriceInfoService02) 기준. 구 경로(ao/PriceInfoService)는 자동 폴백.
키: 환경변수 G2B_KEY (디코딩 인증키)
  $env:G2B_KEY="<키>"; python -X utf8 -u tools\\fetch_g2b_price.py
  python -X utf8 -u tools\\fetch_g2b_price.py --only 종합
  python -X utf8 -u tools\\fetch_g2b_price.py --search "파고라"
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

KEY = os.environ.get("G2B_KEY", "").strip()
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

BASE_CANDIDATES = (
    "http://apis.data.go.kr/1230000/PriceInfoService02/",
    "http://apis.data.go.kr/1230000/ao/PriceInfoService/",
)

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "05_내역서" / "일위대가DB" / "조달청_가격정보_2026"
OUT_DIR.mkdir(parents=True, exist_ok=True)

OPS = [
    ("getPriceInfoListFcltyCmmnMtrilEngrk", "시설공통자재(토목)", "시설공통자재_토목"),
    ("getPriceInfoListFcltyCmmnMtrilBildng", "시설공통자재(건축)", "시설공통자재_건축"),
    ("getPriceInfoListFcltyCmmnMtrilMchnEqp", "시설공통자재(기계설비)", "시설공통자재_기계설비"),
    ("getPriceInfoListFcltyCmmnMtrilElctyIrmc", "시설공통자재(전기·정보통신)", "시설공통자재_전기정보통신"),
    ("getPriceInfoListFcltyCmmnMtrilTotal", "시설공통자재(종합)", "시설공통자재_종합"),
    ("getPriceInfoListMrktCnstrctPcEngrk", "시장시공가격(토목)", "시장시공가격_토목"),
    ("getPriceInfoListMrktCnstrctPcBildng", "시장시공가격(건축)", "시장시공가격_건축"),
    ("getPriceInfoListMrktCnstrctPcMchnEqp", "시장시공가격(기계설비)", "시장시공가격_기계설비"),
    ("getStdMarkUprcinfoList", "표준시장단가", "표준시장단가"),
    ("getCnsttyClsfcInfoList", "공종분류및세부공종", "공종분류"),
    ("getNetRsceinfoList", "자원분류및순수자원", "자원분류_순수자원"),
]

ALT = {
    "getStdMarkUprcinfoList": ["getStdMarkUprcInfoList"],
    "getPriceInfoListFcltyCmmnMtrilTotal": ["getPriceInfoListFcltyCmmnMtrlTotal"],
    "getNetRsceinfoList": ["getNetRsceInfoList"],
    "getCnsttyClsfcInfoList": ["getCnsttyClsfcInfoList"],
}

SEARCH_OPS = {
    "getPriceInfoListFcltyCmmnMtrilEngrk",
    "getPriceInfoListFcltyCmmnMtrilBildng",
    "getPriceInfoListFcltyCmmnMtrilMchnEqp",
    "getPriceInfoListFcltyCmmnMtrilElctyIrmc",
    "getPriceInfoListFcltyCmmnMtrilTotal",
    "getPriceInfoListMrktCnstrctPcEngrk",
    "getPriceInfoListMrktCnstrctPcBildng",
    "getPriceInfoListMrktCnstrctPcMchnEqp",
}

BASE: str | None = None


def call(base: str, op: str, page: int, rows: int = 500, extra: dict | None = None) -> tuple[int, str]:
    params = {
        "serviceKey": KEY,
        "type": "json",
        "numOfRows": str(rows),
        "pageNo": str(page),
    }
    if extra:
        params.update(extra)
    url = base + op + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
        return r.status, r.read().decode("utf-8", "replace")


def parse_json(txt: str) -> dict | None:
    if not txt.lstrip().startswith("{"):
        return None
    try:
        d = json.loads(txt)
    except json.JSONDecodeError:
        return None
    hdr = d.get("response", {}).get("header", {})
    code = str(hdr.get("resultCode", ""))
    if code and code not in ("00", "0"):
        msg = hdr.get("resultMsg", "")
        raise RuntimeError(f"API resultCode={code} {msg}")
    return d


def extract_items(d: dict) -> tuple[int, list[dict]]:
    body = d.get("response", {}).get("body", {})
    total = int(body.get("totalCount") or 0)
    items = body.get("items")
    if items is None:
        return total, []
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return total, list(items or [])


def resolve_base(probe_op: str = "getNetRsceinfoList") -> str:
    global BASE
    if BASE:
        return BASE
    for base in BASE_CANDIDATES:
        for cand in [probe_op] + ALT.get(probe_op, []):
            try:
                st, txt = call(base, cand, 1, rows=3)
                if st == 200 and parse_json(txt) is not None:
                    BASE = base
                    print(f"[BASE] {base}")
                    return base
            except (urllib.error.HTTPError, RuntimeError, OSError):
                continue
    raise SystemExit("[중단] API BASE URL 확인 실패 — G2B_KEY·활용승인·네트워크를 확인하세요.")


def try_op(base: str, op: str) -> tuple[str | None, dict | None]:
    for cand in [op] + ALT.get(op, []):
        try:
            st, txt = call(base, cand, 1, rows=10)
            if st == 200:
                d = parse_json(txt)
                if d is not None:
                    return cand, d
        except (urllib.error.HTTPError, RuntimeError, OSError):
            continue
    return None, None


def search_params(keyword: str | None) -> dict:
    if not keyword:
        return {}
    return {"prdctClsfcNoNm": keyword.strip()}


def fetch_one(base: str, op: str, label: str, fname: str, extra: dict, rows_per: int) -> tuple[str, int]:
    real, d = try_op(base, op)
    if not real:
        print(f"[skip] {label} ({op}) — 404/오류")
        return "오류", 0
    total, first = extract_items(d)
    rows = list(first)
    pages = max(1, (total + rows_per - 1) // rows_per) if total else 1
    for p in range(2, pages + 1):
        try:
            _, txt = call(base, real, p, rows_per, extra)
            _, items = extract_items(parse_json(txt) or {})
            rows.extend(items)
            time.sleep(0.04)
        except Exception as ex:
            print(f"   [경고] {label} p{p} 실패: {ex}")
            break
    if not rows:
        print(f"[빈값] {label} ({real}) total={total}")
        return real, 0
    keys: list[str] = []
    for it in rows:
        for k in it.keys():
            if k not in keys:
                keys.append(k)
    out = OUT_DIR / f"{fname}.csv"
    with out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        for it in rows:
            w.writerow({k: it.get(k, "") for k in keys})
    print(f"[OK]  {label:24s} {real:38s} total={total:>7} saved={len(rows):>7} -> {out.name}")
    return real, len(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="라벨 부분일치 필터(예: 종합, 토목)")
    ap.add_argument("--search", help="품명/규격 부분검색(prdctClsfcNoNm·krnPrdctNm 동시)")
    ap.add_argument("--rows", type=int, default=500, help="페이지당 건수(기본 500)")
    args = ap.parse_args()

    if not KEY:
        print("[중단] 환경변수 G2B_KEY 가 비어 있습니다.")
        sys.exit(1)

    base = resolve_base()
    extra = search_params(args.search)
    if extra:
        print(f"[검색] {args.search!r} (시설공통·시장시공 오퍼레이션에만 적용)")

    targets = OPS
    if args.only:
        targets = [t for t in OPS if args.only in t[1] or args.only in t[2]]
        if not targets:
            raise SystemExit(f"--only {args.only!r} 에 해당하는 오퍼레이션 없음")

    summary = []
    for op, label, fname in targets:
        op_extra = extra if op in SEARCH_OPS else {}
        real, n = fetch_one(base, op, label, fname, op_extra, args.rows)
        summary.append((label, real, n))
        time.sleep(0.05)

    print("\n=== 수집 요약 ===")
    for lab, op, n in summary:
        print(f"  {lab:24s} {n:>7}건  ({op})")
    print(f"\n저장 폴더: {OUT_DIR}")


if __name__ == "__main__":
    main()
