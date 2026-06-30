import os, sys, ssl, json, urllib.parse, urllib.request, time
sys.stdout.reconfigure(encoding="utf-8")
KEY = os.environ.get("G2B_KEY", "").strip()
ctx = ssl.create_default_context(); ctx.check_hostname=False; ctx.verify_mode=ssl.CERT_NONE
def total(nm):
    base={"serviceKey":KEY,"type":"json","numOfRows":"1","pageNo":"1",
          "rgstDtBgnDt":"202507010000","rgstDtEndDt":"202606302359","prdctClsfcNoNm":nm}
    url="https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService/getMASCntrctPrdctInfoList?"+urllib.parse.urlencode(base)
    req=urllib.request.Request(url,headers={"User-Agent":"Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req,timeout=30,context=ctx) as r:
            d=json.loads(r.read().decode("utf-8","replace"))
        if "response" in d: return d["response"]["body"].get("totalCount")
        return "ERR:"+d.get("nkoneps.com.response.ResponseError",{}).get("header",{}).get("resultCode","?")
    except Exception as e:
        return "EXC"
terms=["의자","벤치","장의자","야외탁자","파고라","퍼걸러","그늘막","그늘시렁","정자","쉘터","셸터",
       "운동기구","체력단련","화분","화분대","플랜터","수목보호","보호덮개","트리가드","조경",
       "자전거보관대","펜스","휀스","울타리","평상","목재데크","파라솔","무대","음수대","평의자"]
for t in terms:
    print(f"  {t:10s} {total(t)}")
    time.sleep(0.1)
