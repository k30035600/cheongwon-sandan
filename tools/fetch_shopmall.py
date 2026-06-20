"""조달청_종합쇼핑몰 품목정보 서비스(15129471) → 조경 시설물 계약단가 수집 → CSV.

End Point: https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/getMASCntrctPrdctInfoList
주요 파라미터: rgstDtBgnDt·rgstDtEndDt(등록일시 YYYYMMDDHHMM, 최대 1년 폭), prdctClsfcNoNm(품명 부분검색)
단가 필드: cntrctPrceAmt(계약단가) / 규격 prdctSpecNm / 단위 prdctUnit
키는 환경변수 G2B_KEY 로 주입.
  $env:G2B_KEY="<디코딩키>"; python -X utf8 -u tools\\fetch_shopmall.py
"""
import os, sys, ssl, csv, json, time, urllib.parse, urllib.request, urllib.error
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

KEY = os.environ.get("G2B_KEY", "").strip()
if not KEY:
    print("[중단] 환경변수 G2B_KEY 가 비어 있습니다.")
    sys.exit(1)

ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
URL = "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/getMASCntrctPrdctInfoList"

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "05_내역서" / "일위대가DB" / "조달청_종합쇼핑몰_2026"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# 등록일시 창(YYYYMMDDHHMM) — 검증된 단월 단위로 여러 달 누적(폭 초과 오류 회피)
WINDOWS = [
    ("202505010000", "202505312359"),
    ("202503010000", "202503312359"),
]
PAGES = 3            # 키워드·창당 최대 페이지
ROWS = 20            # 응답이 큰 서비스라 행 수를 낮춰 타임아웃 회피

# 조경 시설물 미매칭 대응 품명(부분검색). 0/오류는 자동 skip.
KEYWORDS = ["퍼걸러", "의자", "벤치", "운동기구", "울타리", "수목보호"]


def call(params, win, tries=2):
    p = {"serviceKey": KEY, "type": "json", "numOfRows": str(ROWS), "pageNo": "1",
         "rgstDtBgnDt": win[0], "rgstDtEndDt": win[1]}
    p.update(params)
    url = URL + "?" + urllib.parse.urlencode(p)
    for t in range(tries):
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
                return r.read().decode("utf-8", "replace")
        except Exception:
            time.sleep(1.0)
    return ""


def parse(txt):
    if not txt.lstrip().startswith("{"):
        return None
    try:
        d = json.loads(txt)
    except Exception:
        return None
    if "response" not in d:
        err = d.get("nkoneps.com.response.ResponseError", {}).get("header", {})
        return {"code": err.get("resultCode"), "items": [], "total": 0}
    body = d["response"]["body"]
    items = body.get("items")
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return {"code": "00", "items": items or [], "total": int(body.get("totalCount") or 0)}


def fetch_keyword(kw):
    rows = []
    codes = set()
    for win in WINDOWS:
        for page in range(1, PAGES + 1):
            r = parse(call({"prdctClsfcNoNm": kw, "pageNo": str(page)}, win))
            if not r:
                codes.add("EXC"); break
            codes.add(str(r["code"]))
            if r["code"] not in ("00", "0") or not r["items"]:
                break
            rows.extend(r["items"])
            if page * ROWS >= r["total"]:
                break
            time.sleep(0.15)
    return rows, codes


def main():
    all_rows, seen = [], set()
    for kw in KEYWORDS:
        rows, codes = fetch_keyword(kw)
        new = 0
        for it in rows:
            key = it.get("prdctIdntNo") or it.get("prdctSpecNm")
            if key in seen:
                continue
            seen.add(key); it["_검색어"] = kw; all_rows.append(it); new += 1
        print(f"  {kw:12s} {len(rows):>4}건 (신규 {new})  codes={sorted(codes)}")

    if not all_rows:
        print("수집 결과 없음 — 키/기간/품명 확인.")
        return
    cols = ["_검색어", "prdctClsfcNoNm", "prdctSpecNm", "prdctUnit", "cntrctPrceAmt",
            "cntrctCorpNm", "prdctClsfcNo", "prdctIdntNo", "prdctLrgclsfcNm",
            "prdctMidclsfcNm", "cntrctMthdNm", "rgstDt"]
    out = OUT_DIR / "종합쇼핑몰_조경시설물.csv"
    with out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for it in all_rows:
            w.writerow({c: it.get(c, "") for c in cols})
    print(f"\n저장: {out}  ({len(all_rows)}건)")


if __name__ == "__main__":
    main()
