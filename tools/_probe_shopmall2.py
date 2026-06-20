import os, sys, ssl, json, urllib.parse, urllib.request, urllib.error
sys.stdout.reconfigure(encoding="utf-8")
KEY = os.environ.get("G2B_KEY", "").strip()
ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE

BASES = [
    "http://apis.data.go.kr/1230000/ao/ShoppingMallPrdctInfoService/",
    "http://apis.data.go.kr/1230000/ShoppingMallPrdctInfoService/",
    "http://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/",
    "https://apis.data.go.kr/1230000/ao/ShoppingMallPrdctInfoService/",
]
OPS = ["getMASCntrctPrdctInfoList", "getUcntrctPrdctInfoList", "getThngPrdnmLocplcAccotShoppmallPrdctInfoList"]
COMMON = {"inqryDiv": "1", "inqryBgnDate": "20260101", "inqryEndDate": "20260131"}

def call(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=25, context=ctx) as r:
            return r.status, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")[:400]
    except Exception as e:
        return -1, str(e)[:200]

for base in BASES:
    for op in OPS:
        for extra in ({}, COMMON):
            p = {"serviceKey": KEY, "type": "json", "numOfRows": "3", "pageNo": "1"}
            p.update(extra)
            url = base + op + "?" + urllib.parse.urlencode(p)
            st, txt = call(url)
            tag = "기간" if extra else "기본"
            snippet = txt.lstrip()[:120].replace("\n", " ")
            print(f"[{st}] {base.split('1230000/')[1]}{op} ({tag}): {snippet}")
