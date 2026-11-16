"""ETF 전종목 종가를 받아 prices.json 으로 저장.
1순위: pykrx (KRX 정보데이터시스템, GitHub Secrets 의 KRX_ID/KRX_PW 로 로그인)
2순위: 네이버 금융 ETF 전종목 (KRX 실패 시)
전종목을 통째로 저장하므로 어떤 종목을 보유했는지는 드러나지 않음."""
import json, os, sys, datetime as dt
import requests

KST = dt.timezone(dt.timedelta(hours=9))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"


def from_krx():
    if not (os.getenv("KRX_ID") and os.getenv("KRX_PW")):
        raise RuntimeError("KRX_ID/KRX_PW 시크릿이 없음")
    from pykrx import stock
    today = dt.datetime.now(KST).date()
    for back in range(10):
        d = (today - dt.timedelta(days=back)).strftime("%Y%m%d")
        df = stock.get_etf_ohlcv_by_ticker(d)
        if df is not None and len(df) and df["종가"].sum() > 0:
            items = {str(t): {"c": int(r["종가"])} for t, r in df.iterrows()}
            return {"source": "KRX", "asof": f"{d[:4]}-{d[4:6]}-{d[6:]}", "items": items}
    raise RuntimeError("KRX: 최근 10일 데이터 없음")


def from_naver():
    r = requests.get(
        "https://finance.naver.com/api/sise/etfItemList.nhn?etfType=0&targetColumn=market_sum&sortOrder=desc",
        headers={"User-Agent": UA, "Referer": "https://finance.naver.com/sise/etfList.naver"}, timeout=20)
    r.raise_for_status()
    text = r.content.decode("utf-8", errors="replace")
    if "�" in text:
        text = r.content.decode("euc-kr", errors="replace")
    items = {x["itemcode"]: {"n": x["itemname"], "c": int(x["nowVal"]), "mc": int(x["marketSum"]) * 10**8}
             for x in json.loads(text)["result"]["etfItemList"]}
    return {"source": "NAVER", "asof": dt.datetime.now(KST).strftime("%Y-%m-%d"), "items": items}


out, errors = None, []
for f in (from_krx, from_naver):
    try:
        out = f()
        break
    except Exception as e:  # 실패하면 다음 경로로
        errors.append(f"{f.__name__}: {e}")
        print("실패 ->", errors[-1])
if not out:
    sys.exit("\n".join(errors))
out["updated"] = dt.datetime.now(dt.timezone.utc).isoformat()
if errors:
    out["note"] = " / ".join(errors)
json.dump(out, open("prices.json", "w", encoding="utf-8"), ensure_ascii=False)
print(out["source"], out["asof"], len(out["items"]), "종목 저장")
