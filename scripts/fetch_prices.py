"""ETF 전종목 종가를 받아 prices.json 으로 저장.
1순위: pykrx (KRX 정보데이터시스템, GitHub Secrets 의 KRX_ID/KRX_PW 로 로그인)
2순위: 네이버 금융 ETF 전종목 (KRX 실패 시)
실행: GitHub Actions (평일 장 마감 후). 전종목을 통째로 저장하므로 어떤 종목을 보유했는지는 드러나지 않음."""
import json, os, sys, datetime as dt
import requests

KST = dt.timezone(dt.timedelta(hours=9))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"


def num(v):
    try:
        return int(float(str(v).replace(",", "")))
    except ValueError:
        return 0


def from_krx():
    if not (os.getenv("KRX_ID") and os.getenv("KRX_PW")):
        raise RuntimeError("KRX_ID/KRX_PW 시크릿이 없음")
    today = dt.datetime.now(KST).date()
    for back in range(10):
        d = (today - dt.timedelta(days=back)).strftime("%Y%m%d")
        items = krx_day(d)
        if items:
            return {"source": "KRX", "asof": f"{d[:4]}-{d[4:6]}-{d[6:]}", "items": items}
    raise RuntimeError("KRX: 최근 10일 데이터 없음")


def krx_day(d):
    """하루치 ETF 전종목: 종목명(n)·종가(c)·시가총액(mc). 휴일이면 빈 값."""
    try:  # KRX '전종목 시세' 원본 (시가총액·종목명 포함)
        from pykrx.website.krx.etx.core import 전종목시세_ETF
        df = 전종목시세_ETF().fetch(d)
        items = {}
        for _, r in df.iterrows():
            c = num(r.get("TDD_CLSPRC"))
            if c:
                items[str(r["ISU_SRT_CD"])] = {"n": str(r.get("ISU_ABBRV", "")).strip(), "c": c, "mc": num(r.get("MKTCAP"))}
        if items:
            return items
    except Exception as e:  # pykrx 내부 구조가 바뀐 경우 아래 공개 함수로
        print("KRX 원본 조회 실패, 기본 함수로 다시 시도:", e)
    from pykrx import stock
    df = stock.get_etf_ohlcv_by_ticker(d)
    if df is not None and len(df) and df["종가"].sum() > 0:
        return {str(t): {"c": int(r["종가"])} for t, r in df.iterrows() if int(r["종가"])}
    return {}


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
