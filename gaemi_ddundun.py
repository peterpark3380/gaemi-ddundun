"""
개미는 뚠뚠 - 종목 밸류에이션 검색기
================================================================
종목코드를 입력하면 KIS Open API에서 데이터를 받아 "자기 과거 평균 대비"
현재 PER/PBR이 싼지 비싼지 판단하는 리포트(HTML)를 만들어준다.
(KOSPI_Valuation_Screener.xlsx 에서 쓴 것과 동일한 방법론: 매달 종가를 그
시점에 이미 공시돼 있었을 연간 EPS/BPS로 나눠 과거 PER/PBR을 재구성하고,
현재 배수가 그 히스토리에서 몇 %ile인지로 판단한다.)

일부 종목(삼성전자/SK하이닉스/삼성전기/LG이노텍)은 관세청 수출입데이터
(customs_export.py로 받은 customs_<label>.json, 같은 폴더에 있어야 함)와의
상관관계 분석도 추가로 붙는다 - 전체기간/2024년이후/2025년이후 상관계수를
자동 계산해서 보여준다.

또한 "매수 타이밍 신호" 섹션이 추가로 붙는다 - 최근 일봉 데이터(약 13개월치)를
받아서 거래량(OBV) 다이버전스와 RSI 다이버전스를 자동으로 탐지한다. 이건 위의
PER/PBR 밸류에이션 판단과는 완전히 별개의 기술적/모멘텀 신호다: 밸류에이션은
"과거 대비 비싸다/싸다"를 말해주고, 이 신호는 "지금 추세가 꺾일 조짐이 있는가"를
말해준다. 둘 다 참고용이며, 특히 다이버전스는 확정적 매수/매도 신호가 아니라
패턴 기반 참고 지표이므로 실제 매매는 손절 기준 등 별도의 리스크 관리와 함께
판단할 것.

HOW TO RUN
  같은 cmd 창에 KIS_APPKEY / KIS_APPSECRET 이 이미 set 되어 있으면:
    python3 gaemi_ddundun.py 011070 LG이노텍
    python3 gaemi_ddundun.py 011070          (종목명 생략 가능)
    python3 gaemi_ddundun.py                 (인자 없이 실행하면 직접 입력받음)

  수출입데이터 상관관계까지 보려면 같은 폴더에 아래 파일들이 있어야 함
  (customs_export.py로 미리 받아둘 것):
    005930(삼성전자)/000660(SK하이닉스)/067310(하나마이크론): customs_dram.json, customs_hbm.json, customs_nand.json
    009150(삼성전기): customs_mlcc.json, customs_camera.json
    011070(LG이노텍): customs_camera.json, customs_pcb_c.json
  (475150 SK이터닉스는 내수 발전·REC 판매가 핵심인 사업모델이라 수출입 품목 매핑을 넣지 않았음)
  파일이 없으면 그 섹션은 자동으로 생략되고 나머지 리포트는 정상 생성된다.

  증권사 목표주가/선행 밸류에이션까지 보려면 같은 폴더에 analyst_targets.json이 있어야 함
  (KIS API로 자동 조회되는 게 아니라, 증권사 리포트를 보고 직접 입력/갱신하는 파일 - 형식은
  파일 안 예시 참고). 파일이 없거나 해당 종목 항목이 없으면 그 섹션은 자동으로 생략된다.

  "과거 급등 이벤트 백테스트"까지 보려면 미리 event_backtest.py를 한 번 실행해서
  event_backtest_<종목코드>.json을 만들어둬야 함:
    python3 event_backtest.py 011070 LG이노텍
  (일봉 데이터를 몇 년치 여러 번 나눠 받아오기 때문에 시간이 좀 걸린다). 실행해두면 이후
  gaemi_ddundun.py 검색 시 자동으로 반영된다. 파일이 없으면 그 섹션은 생략된다.

  "PER-주가 상관관계 분석"(상관계수 + 지수화 차트 + PER 구간별 이후 수익률 백테스트)과
  "외국인·프로그램매매 수급"(당일 외국인 보유율/순매수 + 최근 추이) 섹션은 추가 설정 없이
  자동으로 붙는다 (기존에 이미 받아오던 월별 데이터/현재가 조회 응답을 재활용).

OUTPUT
  개미는뚠뚠_리포트_<종목코드>.html  - 더블클릭하면 브라우저에서 바로 열림
"""

import os
import sys
import json
import time
import statistics
import requests

BASE_URL = "https://openapi.koreainvestment.com:9443"

# 종목코드 -> [(표시이름, customs_<label>.json 의 label), ...]
STOCK_TO_PRODUCTS = {
    "005930": [("DRAM", "dram"), ("HBM/MCP", "hbm"), ("NAND", "nand")],
    "000660": [("DRAM", "dram"), ("HBM/MCP", "hbm"), ("NAND", "nand")],
    "009150": [("MLCC", "mlcc"), ("카메라모듈", "camera")],
    "011070": [("카메라모듈", "camera"), ("PCB(기타)", "pcb_c")],
    "067310": [("DRAM", "dram"), ("HBM/MCP", "hbm"), ("NAND", "nand")],  # 하나마이크론: SK하이닉스 등의
    # 메모리 후공정(패키징/테스트) 외주 물량이 핵심 사업이라, 메모리 수출 물량과 간접적으로 연동될
    # 것으로 보고 기존 DRAM/HBM/NAND 데이터를 그대로 재사용 (SK하이닉스와 동일 카테고리).
    # SK이터닉스(475150, 태양광/풍력/ESS/수소 IPP)는 국내 발전·REC 판매가 핵심인 내수 사업모델이라
    # 관세청 수출입 품목 데이터와 자연스럽게 대응되는 카테고리가 없다고 판단해 넣지 않았음 - 억지로
    # 끼워맞추느니 생략하는 쪽을 택함 (근거 없으면 빈 리스트/미등록이 맞다는 원칙).
}


# --------------------------------------------------------------------- KIS API

def get_credentials():
    appkey = os.environ.get("KIS_APPKEY")
    appsecret = os.environ.get("KIS_APPSECRET")
    if not appkey or not appsecret:
        sys.exit("Missing KIS_APPKEY / KIS_APPSECRET environment variables. "
                  "같은 cmd 창에 set KIS_APPKEY=... / set KIS_APPSECRET=... 먼저 실행해주세요.")
    return appkey, appsecret


TOKEN_CACHE_FILE = "gaemi_ddundun_token_cache.json"


def _hash_key(appkey):
    import hashlib
    return hashlib.sha256(appkey.encode("utf-8")).hexdigest()


def _load_cached_token(appkey):
    """KIS 토큰 발급 API는 1분에 1회로 제한돼 있다 - 여러 종목을 연달아 검색하면 바로 막히므로,
    유효한 토큰이 있으면 재발급 없이 재사용한다 (KIS 토큰 유효기간은 통상 24시간)."""
    if not os.path.exists(TOKEN_CACHE_FILE):
        return None
    try:
        from datetime import datetime, timedelta
        with open(TOKEN_CACHE_FILE, encoding="utf-8") as f:
            cache = json.load(f)
        if cache.get("appkey_hash") != _hash_key(appkey):
            return None
        expires_at = datetime.fromisoformat(cache["expires_at"])
        if datetime.now() < expires_at - timedelta(minutes=5):
            return cache["access_token"]
    except Exception:
        pass
    return None


def _save_token_cache(appkey, access_token, expires_in_sec):
    from datetime import datetime, timedelta
    expires_at = datetime.now() + timedelta(seconds=int(expires_in_sec))
    cache = {"appkey_hash": _hash_key(appkey), "access_token": access_token,
             "expires_at": expires_at.isoformat()}
    try:
        with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        pass  # 캐시 저장 실패해도 리포트 생성 자체는 계속 진행


def get_access_token(appkey, appsecret):
    cached = _load_cached_token(appkey)
    if cached:
        print("(캐시된 토큰 재사용 - 1분에 1회 발급 제한 회피)")
        return cached

    url = f"{BASE_URL}/oauth2/tokenP"
    body = {"grant_type": "client_credentials", "appkey": appkey, "appsecret": appsecret}
    try:
        r = requests.post(url, json=body, timeout=15)
        r.raise_for_status()
    except requests.exceptions.HTTPError:
        if r.status_code == 403:
            sys.exit(
                "토큰 발급이 막혔습니다 (HTTP 403). KIS Open API는 접근토큰 발급을 1분에 1회로 "
                "제한합니다. 방금 다른 종목을 검색해서 토큰을 이미 발급받았다면 60초 정도 "
                "기다렸다가 다시 실행해주세요."
            )
        raise
    data = r.json()
    if "access_token" not in data:
        raise RuntimeError(f"Token issuance failed: {data}")
    _save_token_cache(appkey, data["access_token"], data.get("expires_in", 86400))
    return data["access_token"]


def _is_rate_limited(body):
    return isinstance(body, dict) and body.get("rt_cd") == "1" and body.get("msg_cd") == "EGW00201"


def _call_with_retry(fetch_fn, *args, max_retries=5, **kwargs):
    status, body = None, None
    for attempt in range(max_retries):
        status, body = fetch_fn(*args, **kwargs)
        if _is_rate_limited(body):
            wait = 1.5 * (attempt + 1)
            print(f"    [초당 호출 제한, {wait:.1f}초 후 재시도 {attempt + 1}/{max_retries}]")
            time.sleep(wait)
            continue
        return status, body
    return status, body


def get_current_quote(token, appkey, appsecret, symbol):
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-price"
    headers = {"content-type": "application/json; charset=utf-8", "authorization": f"Bearer {token}",
               "appkey": appkey, "appsecret": appsecret, "tr_id": "FHKST01010100"}
    params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": symbol}
    r = requests.get(url, headers=headers, params=params, timeout=15)
    return r.status_code, r.json()


def get_financial_ratio(token, appkey, appsecret, symbol):
    url = f"{BASE_URL}/uapi/domestic-stock/v1/finance/financial-ratio"
    headers = {"content-type": "application/json; charset=utf-8", "authorization": f"Bearer {token}",
               "appkey": appkey, "appsecret": appsecret, "tr_id": "FHKST66430300"}
    params = {"fid_input_iscd": symbol, "fid_div_cls_code": "0", "fid_cond_mrkt_div_code": "J"}
    r = requests.get(url, headers=headers, params=params, timeout=15)
    return r.status_code, r.json()


def get_price_history(token, appkey, appsecret, symbol, start, end, period="M"):
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    headers = {"content-type": "application/json; charset=utf-8", "authorization": f"Bearer {token}",
               "appkey": appkey, "appsecret": appsecret, "tr_id": "FHKST03010100"}
    params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": symbol, "FID_INPUT_DATE_1": start,
              "FID_INPUT_DATE_2": end, "FID_PERIOD_DIV_CODE": period, "FID_ORG_ADJ_PRC": "0"}
    r = requests.get(url, headers=headers, params=params, timeout=20)
    return r.status_code, r.json()


def get_daily_bars_chunked(token, appkey, appsecret, symbol, lookback_days=400, chunk_days=95):
    """일봉 데이터는 한 번의 호출로 lookback_days 전체를 못 받아올 수 있어서(회신 건수 제한),
    최근 날짜부터 chunk_days 단위로 거슬러 올라가며 여러 번 호출해 합친다.
    반환: [{"date": "YYYYMMDD", "close": float, "volume": float}, ...] 날짜 오름차순, 중복 제거됨."""
    from datetime import datetime, timedelta
    fmt = "%Y%m%d"
    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=lookback_days)
    by_date = {}
    cursor_end = end_dt
    while cursor_end >= start_dt:
        cursor_start = max(start_dt, cursor_end - timedelta(days=chunk_days))
        s_str, e_str = cursor_start.strftime(fmt), cursor_end.strftime(fmt)
        try:
            _, body = _call_with_retry(get_price_history, token, appkey, appsecret, symbol, s_str, e_str, "D")
        except Exception:
            body = None
        if isinstance(body, dict):
            for b in body.get("output2", []):
                d = b.get("stck_bsop_date")
                close = b.get("stck_clpr")
                vol = b.get("acml_vol")
                if d and close not in (None, "") and vol not in (None, ""):
                    by_date[d] = {"date": d, "close": float(close), "volume": float(vol)}
        time.sleep(0.7)
        cursor_end = cursor_start - timedelta(days=1)
    return [by_date[d] for d in sorted(by_date)]


def get_investor_trend(token, appkey, appsecret, symbol):
    """일별 투자자매매동향(최근 거래일 기준 외국인/기관/개인 순매수) - tr_id FHKST01010900.
    주의: 이 엔드포인트는 현재가 조회(FHKST01010100)만큼 실전 검증을 많이 못 해봤다. 응답 스키마가
    예상과 다르면 render_investor_flow_section()에서 그 부분만 조용히 생략된다(리포트 자체는 안 깨짐)."""
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-investor"
    headers = {"content-type": "application/json; charset=utf-8", "authorization": f"Bearer {token}",
               "appkey": appkey, "appsecret": appsecret, "tr_id": "FHKST01010900"}
    params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": symbol}
    r = requests.get(url, headers=headers, params=params, timeout=15)
    return r.status_code, r.json()


# --------------------------------------------------------------------- 방법론

def known_fy(year, month):
    """4~12월엔 직전연도 실적, 1~3월엔 전전연도 실적이 '그 시점에 이미 공시돼 있었을' 값."""
    return year - 1 if month >= 4 else year - 2


def build_history(annual, bars):
    """annual: {year: (eps, bps)}.  bars: KIS output2 list (최신순).  반환: [(ym, price, per, pbr), ...] 오름차순."""
    rows = []
    for b in bars:
        d = b["stck_bsop_date"]
        y, m = int(d[:4]), int(d[4:6])
        price = float(b["stck_clpr"])
        kfy = known_fy(y, m)
        if kfy not in annual:
            continue
        eps, bps = annual[kfy]
        per = price / eps if eps and eps > 0 else None
        pbr = price / bps if bps and bps > 0 else None
        rows.append((f"{y:04d}-{m:02d}", price, per, pbr))
    rows.sort()
    return rows


def percentile_of(value, series):
    valid = [v for v in series if v is not None]
    if not valid or value is None:
        return None
    return sum(1 for v in valid if v <= value) / len(valid)


def verdict_of(pct):
    if pct is None:
        return "판단불가"
    if pct < 0.20:
        return "매우저평가"
    if pct < 0.40:
        return "저평가"
    if pct < 0.60:
        return "적정"
    if pct < 0.80:
        return "고평가"
    return "매우고평가"


VERDICT_COLOR = {
    "매우저평가": ("#0ca30c", "#0ca30c"),
    "저평가": ("#0ca30c", "#0ca30c"),
    "적정": ("#898781", "#898781"),
    "고평가": ("#fab219", "#c98500"),
    "매우고평가": ("#d03b3b", "#e66767"),
    "판단불가": ("#898781", "#898781"),
}


# --------------------------------------------------------------------- 수출입데이터 상관관계
# (customs_export.py 로 미리 받아둔 customs_<label>.json 이 같은 폴더에 있을 때만 동작)

def load_customs_local(label):
    fname = f"customs_{label}.json"
    if not os.path.exists(fname):
        return None
    with open(fname, encoding="utf-8") as f:
        d = json.load(f)
    out = {}
    for it in d.get("items", []):
        if it.get("year") == "총계":
            continue
        y, m = it["year"].split(".")
        ym = f"{y}-{m}"
        try:
            expDlr = float(it["expDlr"])
            expWgt = float(it["expWgt"])
        except (TypeError, ValueError):
            continue
        out[ym] = {"expDlr": expDlr, "expWgt": expWgt,
                   "unit_price": expDlr / expWgt if expWgt > 0 else None}
    return out


def _pearson(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    cov = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    vx = sum((x - mx) ** 2 for x in xs)
    vy = sum((y - my) ** 2 for y in ys)
    if vx <= 0 or vy <= 0:
        return None
    return cov / (vx * vy) ** 0.5


def export_corr_analysis(bars, cdict):
    """bars: KIS 월별 가격 output2.  cdict: load_customs_local() 결과.
    반환: {"수출액": {"전체기간": r, "2024-01~": r, "2025-01~": r}, "물량": {...}, "단가": {...},
           "n": int, "range": (start_ym, end_ym)}  또는 데이터 부족 시 None."""
    price_by_ym = {}
    for b in bars:
        d = b["stck_bsop_date"]
        ym = f"{d[:4]}-{d[4:6]}"
        price_by_ym[ym] = float(b["stck_clpr"])
    common = sorted(set(price_by_ym) & set(cdict))
    if len(common) < 12:
        return None

    def corr_for(metric):
        results = {}
        for label, cutoff in [("전체기간", None), ("2024-01~", "2024-01"), ("2025-01~", "2025-01")]:
            sub = [ym for ym in common if (cutoff is None or ym >= cutoff)]
            pairs = [(price_by_ym[ym], cdict[ym][metric]) for ym in sub if cdict[ym].get(metric) is not None]
            results[label] = _pearson([p[0] for p in pairs], [p[1] for p in pairs]) if len(pairs) >= 6 else None
        return results

    return {
        "수출액": corr_for("expDlr"),
        "물량": corr_for("expWgt"),
        "단가": corr_for("unit_price"),
        "n": len(common),
        "range": (common[0], common[-1]),
    }


def corr_verdict(r):
    if r is None:
        return "데이터 부족", "#898781"
    ar = abs(r)
    if ar >= 0.5:
        return "의미있는 신호", "#0ca30c"
    if ar >= 0.3:
        return "약한 신호", "#fab219"
    return "무의미", "#898781"


# --------------------------------------------------------------------- PER-주가 상관관계 분석
# 주의: PER = 주가 ÷ EPS 이므로, 같은 회계연도 실적(EPS 고정)이 적용되는 구간 안에서는 주가와
# PER이 정의상 거의 기계적으로 함께 움직인다 (분모가 상수인 비례식). 그래서 아래 상관계수는
# "PER이 주가를 예측한다"는 의미가 아니라, 그 시기에 이익(EPS)이 주가보다 빠르게/느리게
# 바뀌었는지를 드러내는 지표에 가깝다 - 리포트에도 이 점을 항상 명시한다. 좀 더 실질적으로
# "밸류에이션이 낮았을 때 실제로 이후 수익률이 더 좋았는가"를 보려면 per_forward_return_analysis()
# (구간별 이후 수익률)가 더 적합하다.

def per_price_correlation(history):
    """history: [(ym, price, per, pbr), ...] 오름차순. 전체기간/최근24개월/최근12개월 Pearson r."""
    valid = [(price, per) for _, price, per, _ in history if per is not None]
    if len(valid) < 12:
        return None

    def corr_for(n_months):
        sub = valid[-n_months:] if n_months else valid
        if len(sub) < 6:
            return None
        return _pearson([p for p, _ in sub], [pe for _, pe in sub])

    return {"전체기간": corr_for(None), "최근24개월": corr_for(24), "최근12개월": corr_for(12), "n": len(valid)}


def _add_months(ym, h):
    y, m = map(int, ym.split("-"))
    total = (y * 12 + (m - 1)) + h
    ny, nm = divmod(total, 12)
    return f"{ny:04d}-{nm + 1:02d}"


def per_forward_return_analysis(history, horizons_months=(6, 12)):
    """PER이 '그 시점까지의 과거'(확장윈도우, 미래데이터 미포함) 중 몇 %ile였는지로 5구간을 나누고,
    구간별로 이후 h개월 뒤 수익률을 집계. 반환: {bucket: {h: {"n","mean","median","win_rate"}}}
    또는 데이터 부족 시 None."""
    valid = [row for row in history if row[2] is not None]
    if len(valid) < 24:
        return None
    by_ym = {row[0]: row for row in history}

    def classify(pct):
        if pct < 0.20:
            return "0-20%(저평가)"
        if pct < 0.40:
            return "20-40%"
        if pct < 0.60:
            return "40-60%"
        if pct < 0.80:
            return "60-80%"
        return "80-100%(고평가)"

    buckets = {b: {h: [] for h in horizons_months} for b in
               ["0-20%(저평가)", "20-40%", "40-60%", "60-80%", "80-100%(고평가)"]}

    for pos, row in enumerate(valid):
        if pos < 12:  # 최소 12개월 이상 쌓인 뒤부터 (초기 percentile은 표본이 너무 적어 불안정)
            continue
        ym, price, per, pbr = row
        past_pers = [r[2] for r in valid[:pos + 1]]
        pct = sum(1 for v in past_pers if v <= per) / len(past_pers)
        bucket = classify(pct)
        for h in horizons_months:
            target = by_ym.get(_add_months(ym, h))
            if target is not None and price:
                buckets[bucket][h].append((target[1] - price) / price)

    aggregate = {}
    for b, hmap in buckets.items():
        aggregate[b] = {}
        for h, vals in hmap.items():
            if vals:
                aggregate[b][h] = {"n": len(vals), "mean": statistics.mean(vals),
                                    "median": statistics.median(vals),
                                    "win_rate": sum(1 for v in vals if v > 0) / len(vals)}
            else:
                aggregate[b][h] = {"n": 0, "mean": None, "median": None, "win_rate": None}
    return aggregate


def render_per_price_correlation_section(history):
    corr = per_price_correlation(history)
    fwd = per_forward_return_analysis(history)
    if corr is None and fwd is None:
        return ""

    corr_html = ""
    if corr is not None:
        rows = []
        for label in ["전체기간", "최근24개월", "최근12개월"]:
            r = corr[label]
            vtext, vcolor = corr_verdict(r)
            rows.append(f"<tr><td>{label}</td><td>{fmt_r(r)}</td>"
                        f"<td><span class='corr-badge' style='background:{vcolor}'>{vtext}</span></td></tr>")
        chart = index_dual_chart_svg(history)
        corr_html = f"""
    <div class="tech-subtitle">PER-주가 상관계수 (Pearson r)</div>
    <table class="corr-table">
      <tr><th>구간</th><th>r</th><th>판정</th></tr>
      {''.join(rows)}
    </table>
    <div class="tech-subtitle">주가 vs PER 추이 비교 (각각 시작월=100 기준 지수화)</div>
    <div class="chart-legend">
      <span><span class="dot" style="background:var(--series-1)"></span>주가(지수화)</span>
      <span><span class="dot" style="background:var(--series-2)"></span>PER(지수화)</span>
    </div>
    {chart}
    <div class="tech-explain">PER = 주가 ÷ EPS라서, 같은 회계연도 실적이 적용되는 구간(보통 12개월) 안에서는
    주가와 PER이 정의상 거의 기계적으로 함께 움직입니다. 그래서 이 상관계수 자체가 "PER이 주가를 예측한다"는
    뜻은 아니고, 두 선이 벌어지는 시점(예: 주가는 오르는데 PER은 안 오르거나 내리는 구간)이 있다면 그건
    실적(EPS)이 주가보다 더 빠르게 성장하고 있었다는 뜻이라 오히려 눈여겨볼 만합니다.</div>"""

    fwd_html = ""
    if fwd is not None:
        blocks = []
        for h, hlabel in [(6, "6개월 후"), (12, "12개월 후")]:
            rows = []
            for b in ["0-20%(저평가)", "20-40%", "40-60%", "60-80%", "80-100%(고평가)"]:
                a = fwd[b][h]
                win = fmt_pct(a["win_rate"]) if a["win_rate"] is not None else "-"
                rows.append(f"<tr><td>{b}</td><td>{a['n']}</td>"
                            f"<td>{fmt_signed_pct(a['mean'])}</td><td>{fmt_signed_pct(a['median'])}</td><td>{win}</td></tr>")
            blocks.append(f"""
    <div class="tech-subtitle">PER 구간(그 시점까지 과거 대비 %ile)별 {hlabel} 수익률</div>
    <table class="corr-table">
      <tr><th>PER 구간</th><th>표본</th><th>평균수익률</th><th>중앙값</th><th>승률</th></tr>
      {''.join(rows)}
    </table>""")
        fwd_html = f"""
    {''.join(blocks)}
    <div class="tech-explain">PER 백분위는 그 시점까지의 과거 데이터만 사용해 계산했습니다(미래 데이터 미포함,
    사후편향 방지). "0-20%(저평가)" 구간에서 이후 수익률이 실제로 더 좋고 승률도 높다면, 위 밸류에이션
    판단이 이 종목에서는 통계적으로도 근거가 있다는 뜻이고, 반대로 구간별 차이가 뚜렷하지 않다면 이
    종목은 "싸다고 꼭 오르지는 않는" 유형일 수 있습니다.</div>"""

    if not corr_html and not fwd_html:
        return ""

    return f"""
  <div class="card">
    <h2>PER-주가 상관관계 분석</h2>
    {corr_html}
    {fwd_html}
  </div>"""


# --------------------------------------------------------------------- 매수 타이밍 신호 (기술적 분석)
# 밸류에이션(PER/PBR)과는 별개의 신호. 최근 일봉 데이터에서 가격의 스윙 저점/고점을 찾고,
# 그 지점에서 거래량(OBV)과 모멘텀(RSI)이 가격과 같은 방향인지 반대 방향인지(다이버전스)를 본다.
# - 강세(불리시) 다이버전스: 가격은 저점을 낮추는데 OBV/RSI는 저점을 높임 -> 매도 압력 약화 신호
# - 약세(베어리시) 다이버전스: 가격은 고점을 높이는데 OBV/RSI는 고점을 낮춤 -> 상승 동력 약화 신호
# 확정적 매매 신호가 아니라 패턴 기반 참고 지표임을 리포트에 항상 명시한다.

def compute_rsi(closes, period=14):
    """Wilder's RSI. closes: 날짜 오름차순 종가 리스트. 반환: 같은 길이의 리스트, 앞쪽 period개는 None."""
    n = len(closes)
    rsis = [None] * n
    if n <= period:
        return rsis
    gains = [max(closes[i] - closes[i - 1], 0) for i in range(1, n)]
    losses = [max(closes[i - 1] - closes[i], 0) for i in range(1, n)]
    avg_gain = sum(gains[:period]) / period
    avg_loss = sum(losses[:period]) / period
    rsis[period] = 100.0 if avg_loss == 0 else 100 - 100 / (1 + avg_gain / avg_loss)
    for i in range(period + 1, n):
        avg_gain = (avg_gain * (period - 1) + gains[i - 1]) / period
        avg_loss = (avg_loss * (period - 1) + losses[i - 1]) / period
        rsis[i] = 100.0 if avg_loss == 0 else 100 - 100 / (1 + avg_gain / avg_loss)
    return rsis


def compute_obv(closes, volumes):
    """On-Balance Volume: 종가가 오른 날은 거래량을 더하고 내린 날은 뺀다 (누적)."""
    obv = [0.0] * len(closes)
    for i in range(1, len(closes)):
        if closes[i] > closes[i - 1]:
            obv[i] = obv[i - 1] + volumes[i]
        elif closes[i] < closes[i - 1]:
            obv[i] = obv[i - 1] - volumes[i]
        else:
            obv[i] = obv[i - 1]
    return obv


def moving_average(values, period):
    if len(values) < period:
        return None
    return sum(values[-period:]) / period


def find_swing_points(values, order=5):
    """values 안에서 앞뒤 order개 구간 내 국소 최소/최대인 지점을 스윙 저점/고점으로 판단.
    반환: [(index, value, "low"|"high"), ...] 인덱스 오름차순."""
    n = len(values)
    swings = []
    for i in range(order, n - order):
        window = values[i - order:i + order + 1]
        if values[i] == min(window) and window.count(values[i]) == 1:
            swings.append((i, values[i], "low"))
        elif values[i] == max(window) and window.count(values[i]) == 1:
            swings.append((i, values[i], "high"))
    return swings


def find_trailing_extreme(values, window=15):
    """마지막 window개 구간에서의 최저/최고점 - find_swing_points와 달리 '이후' 구간의
    확인(반등/반락)을 요구하지 않는다. 아직 반등하지 않은 방금 만든 저점/고점을 놓치지
    않기 위한 보조 판단(스윙 저점은 정의상 '반등해야' 국소최소로 확인되므로, 반등 전날까지의
    데이터만 보는 상황 - 다이버전스 사전탐지, 사건 직전 시점 등 - 에서는 find_swing_points만으로
    가장 최근/가장 중요한 저점·고점을 아예 못 찾는 경우가 생긴다)."""
    n = len(values)
    start = max(0, n - window)
    seg = values[start:]
    min_v, max_v = min(seg), max(seg)
    min_i = start + seg.index(min_v)
    max_i = start + seg.index(max_v)
    return (min_i, min_v), (max_i, max_v)


def detect_divergence(dates, closes, obv, rsi, order=5, trailing_window=15):
    """가장 최근의 유의미한 스윙 저점 쌍 또는 고점 쌍(둘 중 더 최근에 발생한 쪽)을 비교해
    가격 vs 거래량(OBV)/모멘텀(RSI) 다이버전스 여부를 판단."""
    swings = find_swing_points(closes, order=order)
    lows = [s for s in swings if s[2] == "low"]
    highs = [s for s in swings if s[2] == "high"]

    candidates = []
    if len(lows) >= 2:
        candidates.append(("bullish_check", lows[-2], lows[-1]))
    if len(highs) >= 2:
        candidates.append(("bearish_check", highs[-2], highs[-1]))

    # 아직 반등/반락으로 '확인'되지 않은 최근 저점/고점도 후보로 포함 (위 설명 참고) -
    # 확인된 마지막 스윙과 비교해서, 반등 전이라도 다이버전스를 놓치지 않는다.
    if len(closes) >= trailing_window:
        (tmin_i, tmin_v), (tmax_i, tmax_v) = find_trailing_extreme(closes, trailing_window)
        if lows and tmin_i > lows[-1][0] and tmin_v < lows[-1][1]:
            candidates.append(("bullish_check", lows[-1], (tmin_i, tmin_v, "low")))
        if highs and tmax_i > highs[-1][0] and tmax_v > highs[-1][1]:
            candidates.append(("bearish_check", highs[-1], (tmax_i, tmax_v, "high")))

    if not candidates:
        return {"type": None, "signals": {}, "points": None}

    candidates.sort(key=lambda c: c[2][0])  # 더 최근에 발생한(인덱스 큰) 스윙 쌍 우선
    kind, p1, p2 = candidates[-1]
    i1, price1, _ = p1
    i2, price2, _ = p2

    result = {"type": None, "signals": {"obv": None, "rsi": None}, "points": {
        "date1": dates[i1], "price1": price1, "obv1": obv[i1], "rsi1": rsi[i1],
        "date2": dates[i2], "price2": price2, "obv2": obv[i2], "rsi2": rsi[i2],
    }}

    if kind == "bullish_check" and price2 < price1:
        result["type"] = "bullish"
        result["signals"]["obv"] = obv[i2] > obv[i1]
        result["signals"]["rsi"] = rsi[i2] is not None and rsi[i1] is not None and rsi[i2] > rsi[i1]
    elif kind == "bearish_check" and price2 > price1:
        result["type"] = "bearish"
        result["signals"]["obv"] = obv[i2] < obv[i1]
        result["signals"]["rsi"] = rsi[i2] is not None and rsi[i1] is not None and rsi[i2] < rsi[i1]

    return result


def divergence_verdict(result):
    """(제목, 색상, 설명문) 반환."""
    if result["type"] is None:
        return ("다이버전스 신호 없음", "#898781",
                "최근 데이터에서 비교할 만한 스윙 저점/고점 쌍을 찾지 못했습니다 (뚜렷한 추세 전환 지점이 아직 없음).")

    obv_sig, rsi_sig = result["signals"]["obv"], result["signals"]["rsi"]
    n_confirm = sum(1 for s in (obv_sig, rsi_sig) if s)

    if result["type"] == "bullish":
        if n_confirm == 2:
            return ("강세 다이버전스 포착 (거래량+모멘텀 동반)", "#0ca30c",
                    "주가는 직전 저점보다 더 낮은 저점을 만들었지만, 거래량(OBV)과 RSI는 오히려 저점을 "
                    "높였습니다. 가격 하락에 비해 매도 압력(거래량)과 하락 모멘텀이 약해지고 있다는 뜻으로, "
                    "저가 매수를 고려해볼 만한 구간이라는 신호입니다.")
        if n_confirm == 1:
            which = "거래량(OBV)" if obv_sig else "RSI"
            return (f"약한 강세 다이버전스 ({which}만 확인)", "#fab219",
                    f"주가는 저점을 낮췄지만 {which} 지표만 저점을 높였습니다. 두 지표가 함께 확인돼야 "
                    "신호가 더 신뢰할 만한데, 지금은 하나만 확인돼 강도가 약합니다.")
        return ("다이버전스 없음 (하락 동반 확인)", "#898781",
                "주가와 거래량(OBV)·RSI가 함께 저점을 낮추고 있어, 아직 매도 압력이 이어지고 있는 것으로 "
                "보입니다. 저가매수 타이밍으로 보기엔 이릅니다.")

    # bearish
    if n_confirm == 2:
        return ("약세 다이버전스 포착 (거래량+모멘텀 동반)", "#d03b3b",
                "주가는 직전 고점보다 더 높은 고점을 만들었지만, 거래량(OBV)과 RSI는 오히려 고점을 "
                "낮췄습니다. 가격 상승에 비해 매수 강도가 약해지고 있다는 뜻으로, 상승 동력 둔화 신호입니다.")
    if n_confirm == 1:
        which = "거래량(OBV)" if obv_sig else "RSI"
        return (f"약한 약세 다이버전스 ({which}만 확인)", "#fab219",
                f"주가는 고점을 높였지만 {which} 지표만 고점을 낮췄습니다. 신호 강도는 약합니다.")
    return ("다이버전스 없음 (상승 동반 확인)", "#0ca30c",
            "주가와 거래량(OBV)·RSI가 함께 고점을 높이고 있어, 상승 추세가 아직 건강해 보입니다.")


# --------------------------------------------------------------------- HTML/SVG

def gauge_svg(pct, verdict, width=520, height=64):
    if pct is None:
        pct = 0.5
    light, dark = VERDICT_COLOR[verdict]
    zones = [(0.0, 0.20, "매우저평가"), (0.20, 0.40, "저평가"), (0.40, 0.60, "적정"),
              (0.60, 0.80, "고평가"), (0.80, 1.0, "매우고평가")]
    zone_colors = {"매우저평가": "#0ca30c", "저평가": "#5fbf5f", "적정": "#c3c2b7",
                    "고평가": "#fab219", "매우고평가": "#d03b3b"}
    bar_y, bar_h = 24, 16
    parts = [f'<svg class="gauge" viewBox="0 0 {width} {height}" width="100%" height="{height}" role="img" aria-label="백분위 {pct*100:.0f}%">']
    for s, e, label in zones:
        x1, x2 = s * width, e * width
        parts.append(f'<rect x="{x1:.1f}" y="{bar_y}" width="{(x2-x1):.1f}" height="{bar_h}" fill="{zone_colors[label]}" opacity="0.85"/>')
    marker_x = pct * width
    parts.append(f'<line x1="{marker_x:.1f}" y1="{bar_y-6}" x2="{marker_x:.1f}" y2="{bar_y+bar_h+6}" stroke="var(--text-primary)" stroke-width="3"/>')
    parts.append(f'<circle cx="{marker_x:.1f}" cy="{bar_y-10}" r="5" fill="var(--text-primary)"/>')
    parts.append(f'<text x="{marker_x:.1f}" y="{bar_y-16}" text-anchor="middle" class="gauge-label">{pct*100:.0f}%ile</text>')
    parts.append(f'<text x="0" y="{bar_y+bar_h+22}" class="gauge-axis">0% (가장 쌈)</text>')
    parts.append(f'<text x="{width}" y="{bar_y+bar_h+22}" text-anchor="end" class="gauge-axis">100% (가장 비쌈)</text>')
    parts.append('</svg>')
    return "".join(parts)


def line_chart_svg(history, key_idx, label, width=760, height=220):
    """key_idx: 2=PER, 3=PBR 컬럼 인덱스"""
    vals = [(i, row[key_idx]) for i, row in enumerate(history) if row[key_idx] is not None]
    if len(vals) < 3:
        return "<p class='muted'>차트를 그리기엔 데이터가 부족합니다.</p>"
    ys = [v for _, v in vals]
    avg = statistics.mean(ys)
    y_min, y_max = min(ys) * 0.95, max(ys) * 1.05
    n = len(history)
    SM = {"left": 50, "right": 20, "top": 16, "bottom": 30}
    plot_w = width - SM["left"] - SM["right"]
    plot_h = height - SM["top"] - SM["bottom"]

    def lx(i):
        return SM["left"] + (i / max(n - 1, 1)) * plot_w

    def ly(v):
        return SM["top"] + plot_h - (v - y_min) / (y_max - y_min) * plot_h

    pts = " ".join(f"{lx(i):.1f},{ly(v):.1f}" for i, v in vals)
    avg_y = ly(avg)
    last_i, last_v = vals[-1]
    parts = [f'<svg class="linechart" viewBox="0 0 {width} {height}" width="100%" height="{height}">']
    # gridlines
    for frac in (0, 0.25, 0.5, 0.75, 1.0):
        gy = SM["top"] + plot_h * frac
        parts.append(f'<line x1="{SM["left"]}" y1="{gy:.1f}" x2="{width-SM["right"]}" y2="{gy:.1f}" class="grid"/>')
    parts.append(f'<line x1="{SM["left"]}" y1="{avg_y:.1f}" x2="{width-SM["right"]}" y2="{avg_y:.1f}" class="avgline" stroke-dasharray="4,4"/>')
    parts.append(f'<text x="{width-SM["right"]}" y="{avg_y-4:.1f}" text-anchor="end" class="avg-label">평균 {avg:.1f}</text>')
    parts.append(f'<polyline points="{pts}" fill="none" class="series-line"/>')
    parts.append(f'<circle cx="{lx(last_i):.1f}" cy="{ly(last_v):.1f}" r="4.5" class="series-dot"/>')
    parts.append(f'<text x="{lx(last_i):.1f}" y="{ly(last_v)-10:.1f}" text-anchor="end" class="cur-label">현재 {last_v:.1f}</text>')
    parts.append(f'<text x="{SM["left"]}" y="{height-6}" class="axis-label">{history[0][0]}</text>')
    parts.append(f'<text x="{width-SM["right"]}" y="{height-6}" text-anchor="end" class="axis-label">{history[-1][0]}</text>')
    parts.append('</svg>')
    return "".join(parts)


def index_dual_chart_svg(history, width=760, height=240):
    """history: [(ym, price, per, pbr), ...]. 주가와 PER을 각각 첫 시점=100으로 지수화해서
    같은 축(0~) 위에 두 선으로 겹쳐 그린다 (듀얼축 대신 지수화 - dataviz 규칙 준수)."""
    valid = [(i, price, per) for i, (ym, price, per, pbr) in enumerate(history) if per is not None]
    if len(valid) < 3:
        return "<p class='muted'>차트를 그리기엔 데이터가 부족합니다.</p>"
    base_price = valid[0][1]
    base_per = valid[0][2]
    if not base_price or not base_per:
        return "<p class='muted'>차트를 그리기엔 데이터가 부족합니다.</p>"
    price_idx = [(i, price / base_price * 100) for i, price, per in valid]
    per_idx = [(i, per / base_per * 100) for i, price, per in valid]
    ys_all = [v for _, v in price_idx] + [v for _, v in per_idx]
    y_min, y_max = min(ys_all) * 0.95, max(ys_all) * 1.05
    n = len(history)
    SM = {"left": 46, "right": 20, "top": 16, "bottom": 30}
    plot_w = width - SM["left"] - SM["right"]
    plot_h = height - SM["top"] - SM["bottom"]

    def lx(i):
        return SM["left"] + (i / max(n - 1, 1)) * plot_w

    def ly(v):
        return SM["top"] + plot_h - (v - y_min) / (y_max - y_min) * plot_h

    price_pts = " ".join(f"{lx(i):.1f},{ly(v):.1f}" for i, v in price_idx)
    per_pts = " ".join(f"{lx(i):.1f},{ly(v):.1f}" for i, v in per_idx)
    base_y = ly(100)
    parts = [f'<svg class="linechart" viewBox="0 0 {width} {height}" width="100%" height="{height}">']
    for frac in (0, 0.25, 0.5, 0.75, 1.0):
        gy = SM["top"] + plot_h * frac
        parts.append(f'<line x1="{SM["left"]}" y1="{gy:.1f}" x2="{width-SM["right"]}" y2="{gy:.1f}" class="grid"/>')
    parts.append(f'<line x1="{SM["left"]}" y1="{base_y:.1f}" x2="{width-SM["right"]}" y2="{base_y:.1f}" class="avgline" stroke-dasharray="4,4"/>')
    parts.append(f'<polyline points="{price_pts}" fill="none" class="series-line"/>')
    parts.append(f'<polyline points="{per_pts}" fill="none" class="series-line-2"/>')
    parts.append(f'<text x="{SM["left"]}" y="{height-6}" class="axis-label">{history[0][0]}</text>')
    parts.append(f'<text x="{width-SM["right"]}" y="{height-6}" text-anchor="end" class="axis-label">{history[-1][0]}</text>')
    parts.append('</svg>')
    return "".join(parts)


def swing_chart_svg(dates, values, mark_indices, title_fmt, width=760, height=160):
    """일별 시계열(가격 또는 OBV)을 그리고, mark_indices의 두 지점(이전/최근 스윙)을 점으로 표시."""
    n = len(values)
    if n < 3:
        return "<p class='muted'>차트를 그리기엔 데이터가 부족합니다.</p>"
    y_min, y_max = min(values), max(values)
    pad = (y_max - y_min) * 0.1 or abs(y_max) * 0.1 or 1
    y_min, y_max = y_min - pad, y_max + pad
    SM = {"left": 60, "right": 20, "top": 16, "bottom": 22}
    plot_w = width - SM["left"] - SM["right"]
    plot_h = height - SM["top"] - SM["bottom"]

    def lx(i):
        return SM["left"] + (i / max(n - 1, 1)) * plot_w

    def ly(v):
        return SM["top"] + plot_h - (v - y_min) / (y_max - y_min) * plot_h

    pts = " ".join(f"{lx(i):.1f},{ly(v):.1f}" for i, v in enumerate(values))
    parts = [f'<svg class="linechart" viewBox="0 0 {width} {height}" width="100%" height="{height}">']
    for frac in (0, 0.5, 1.0):
        gy = SM["top"] + plot_h * frac
        parts.append(f'<line x1="{SM["left"]}" y1="{gy:.1f}" x2="{width-SM["right"]}" y2="{gy:.1f}" class="grid"/>')
    parts.append(f'<polyline points="{pts}" fill="none" class="series-line"/>')
    mark_colors = ["#898781", "#2a78d6"]
    for k, i in enumerate(mark_indices):
        if i is None or i >= n:
            continue
        cx, cy = lx(i), ly(values[i])
        color = mark_colors[min(k, len(mark_colors) - 1)]
        parts.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="5" fill="{color}" stroke="var(--surface-1)" stroke-width="2"/>')
        label = title_fmt(dates[i], values[i])
        anchor = "start" if k == 0 else "end"
        dx = 8 if k == 0 else -8
        parts.append(f'<text x="{cx+dx:.1f}" y="{cy-10:.1f}" text-anchor="{anchor}" class="cur-label" style="fill:{color}">{label}</text>')
    parts.append(f'<text x="{SM["left"]}" y="{height-6}" class="axis-label">{dates[0]}</text>')
    parts.append(f'<text x="{width-SM["right"]}" y="{height-6}" text-anchor="end" class="axis-label">{dates[-1]}</text>')
    parts.append('</svg>')
    return "".join(parts)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>개미는 뚠뚠 - {name} 밸류에이션 리포트</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<style>
  .viz-root {{
    color-scheme: light;
    --surface-1: #fcfcfb; --page: #f9f9f7; --text-primary: #0b0b0b; --text-secondary: #52514e;
    --muted: #898781; --grid: #e1e0d9; --series-1: #2a78d6; --series-2: #eb6834; --border: rgba(11,11,11,0.10);
  }}
  @media (prefers-color-scheme: dark) {{
    :root:where(:not([data-theme="light"])) .viz-root {{
      color-scheme: dark;
      --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff; --text-secondary: #c3c2b7;
      --muted: #898781; --grid: #2c2c2a; --series-1: #3987e5; --series-2: #d95926; --border: rgba(255,255,255,0.10);
    }}
  }}
  :root[data-theme="dark"] .viz-root {{
    color-scheme: dark;
    --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff; --text-secondary: #c3c2b7;
    --muted: #898781; --grid: #2c2c2a; --series-1: #3987e5; --series-2: #d95926; --border: rgba(255,255,255,0.10);
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family: system-ui,-apple-system,"Segoe UI",sans-serif; background: var(--page); color: var(--text-primary); }}
  .viz-root {{ max-width: 860px; margin: 0 auto; padding: 32px 24px 48px; }}
  h1 {{ font-size: 22px; margin: 0 0 4px; }}
  .subtitle {{ color: var(--text-secondary); font-size: 13px; margin-bottom: 24px; }}
  .card {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 12px; padding: 20px 24px; margin-bottom: 16px; }}
  .card h2 {{ font-size: 15px; margin: 0 0 14px; color: var(--text-secondary); font-weight: 600; }}
  .verdict-row {{ display:flex; align-items:center; gap:14px; margin-bottom: 6px; }}
  .verdict-badge {{ display:inline-block; padding: 6px 16px; border-radius: 999px; font-weight:700; font-size: 15px; color:#fff; }}
  .stat-grid {{ display:grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin: 4px 0 6px; }}
  .stat-tile {{ background: var(--page); border-radius: 8px; padding: 10px 12px; }}
  .stat-tile .label {{ font-size: 11px; color: var(--muted); margin-bottom:4px; }}
  .stat-tile .value {{ font-size: 18px; font-weight: 700; font-variant-numeric: tabular-nums; }}
  .gauge-block {{ margin: 18px 0; }}
  .gauge-title {{ font-size: 13px; color: var(--text-secondary); margin-bottom: 6px; }}
  .gauge-label {{ font-size: 11px; font-weight:700; fill: var(--text-primary); }}
  .gauge-axis {{ font-size: 10px; fill: var(--muted); }}
  .grid {{ stroke: var(--grid); stroke-width: 1; }}
  .avgline {{ stroke: var(--muted); stroke-width: 1.5; }}
  .avg-label {{ font-size: 10px; fill: var(--muted); }}
  .series-line {{ stroke: var(--series-1); stroke-width: 2; }}
  .series-line-2 {{ stroke: var(--series-2); stroke-width: 2; }}
  .series-dot {{ fill: var(--series-1); stroke: var(--surface-1); stroke-width: 2; }}
  .cur-label {{ font-size: 11px; font-weight:700; fill: var(--series-1); }}
  .chart-legend {{ display:flex; gap:16px; font-size: 11.5px; color: var(--text-secondary); margin: 6px 0 4px; }}
  .chart-legend .dot {{ display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:5px; vertical-align:middle; }}
  .axis-label {{ font-size: 10px; fill: var(--muted); }}
  .muted {{ color: var(--muted); font-size: 13px; }}
  .notes {{ font-size: 12.5px; color: var(--text-secondary); line-height: 1.7; }}
  .notes li {{ margin-bottom: 4px; }}
  footer {{ text-align:center; color: var(--muted); font-size: 11px; margin-top: 24px; }}
  .corr-product {{ margin-bottom: 18px; }}
  .corr-product:last-child {{ margin-bottom: 0; }}
  .corr-product h3 {{ font-size: 14px; margin: 0 0 8px; }}
  .corr-table {{ width:100%; border-collapse: collapse; font-size: 12.5px; }}
  .corr-table th, .corr-table td {{ text-align:right; padding: 6px 8px; border-bottom: 1px solid var(--grid); font-variant-numeric: tabular-nums; }}
  .corr-table th:first-child, .corr-table td:first-child {{ text-align:left; }}
  .corr-table th {{ color: var(--muted); font-weight: 600; font-size: 11px; }}
  .corr-badge {{ display:inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight:700; color:#fff; }}
  .corr-range {{ color: var(--muted); font-size: 11px; margin: -2px 0 8px; }}
  .tech-explain {{ font-size: 12.5px; color: var(--text-secondary); line-height: 1.6; margin: 8px 0 16px; }}
  .tech-legend {{ font-size: 11px; color: var(--muted); margin: 2px 0 10px; }}
  .tech-legend .dot {{ display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:4px; vertical-align:middle; }}
  .tech-subtitle {{ font-size: 12.5px; color: var(--text-secondary); font-weight:600; margin: 18px 0 6px; }}
  @media (max-width: 560px) {{
    .viz-root {{ padding: 20px 14px 36px; }}
    h1 {{ font-size: 19px; }}
    .card {{ padding: 16px 14px; }}
    .stat-grid {{ grid-template-columns: repeat(2,1fr); gap: 10px; }}
    .stat-tile .value {{ font-size: 16px; }}
    .corr-table {{ font-size: 11.5px; }}
    .corr-table th, .corr-table td {{ padding: 5px 4px; }}
    .verdict-badge {{ font-size: 13px; padding: 5px 12px; }}
  }}
</style>
</head>
<body>
<div class="viz-root">
  <h1>🐜 개미는 뚠뚠 — {name}({code}) 밸류에이션 리포트</h1>
  <div class="subtitle">데이터 기준: {snapshot_date} (KIS Open API 스냅샷) · 방법론: 자기 과거 평균 대비 (동종업계 비교 아님)</div>

  <div class="card">
    <h2>종합 판단</h2>
    <div class="verdict-row">
      <span class="verdict-badge" style="background:{verdict_color}">{verdict}</span>
      <span class="muted">종합 백분위 {composite_pct}</span>
    </div>
    <div class="stat-grid">
      <div class="stat-tile"><div class="label">현재가</div><div class="value">{price:,.0f}원</div></div>
      <div class="stat-tile"><div class="label">현재 PER</div><div class="value">{cur_per}</div></div>
      <div class="stat-tile"><div class="label">현재 PBR</div><div class="value">{cur_pbr}</div></div>
      <div class="stat-tile"><div class="label">데이터 개월수</div><div class="value">{n_months}개월</div></div>
    </div>
  </div>

  {investor_flow_section}

  <div class="card">
    <h2>PER (주가수익비율) — 과거 {n_months}개월 대비</h2>
    <div class="stat-grid" style="grid-template-columns: repeat(3,1fr);">
      <div class="stat-tile"><div class="label">현재 PER</div><div class="value">{cur_per}</div></div>
      <div class="stat-tile"><div class="label">과거 평균 PER</div><div class="value">{avg_per}</div></div>
      <div class="stat-tile"><div class="label">괴리율</div><div class="value">{per_gap}</div></div>
    </div>
    <div class="gauge-block">
      <div class="gauge-title">현재 PER의 과거 히스토리 내 위치</div>
      {per_gauge}
    </div>
    {per_chart}
  </div>

  {per_price_corr_section}

  <div class="card">
    <h2>PBR (주가순자산비율) — 과거 {n_months}개월 대비</h2>
    <div class="stat-grid" style="grid-template-columns: repeat(3,1fr);">
      <div class="stat-tile"><div class="label">현재 PBR</div><div class="value">{cur_pbr}</div></div>
      <div class="stat-tile"><div class="label">과거 평균 PBR</div><div class="value">{avg_pbr}</div></div>
      <div class="stat-tile"><div class="label">괴리율</div><div class="value">{pbr_gap}</div></div>
    </div>
    <div class="gauge-block">
      <div class="gauge-title">현재 PBR의 과거 히스토리 내 위치</div>
      {pbr_gauge}
    </div>
    {pbr_chart}
  </div>

  {analyst_section}

  {technical_section}

  {event_backtest_section}

  {export_corr_section}

  <div class="card">
    <h2>방법론 &amp; 주의사항</h2>
    <ul class="notes">
      <li>매달 종가 ÷ "그 시점에 이미 시장에 공시돼 있었을 연간 EPS/BPS"로 과거 PER/PBR을 재구성합니다
        (4~12월은 직전연도 실적, 1~3월은 전전연도 실적 - 사업보고서 공시 시차 반영). KIS 실시간 PER/PBR과
        같은 계산 방식입니다.</li>
      <li>백분위 = 현재 PER/PBR이 위 방식으로 재구성한 과거 {n_months}개월 중 몇 %ile에 위치하는지.
        낮을수록(0%에 가까울수록) 과거 대비 싸다는 뜻입니다.</li>
      <li>동종업계/경쟁사 비교가 아니라 "자기 자신의 과거"와 비교하는 방식입니다.</li>
      <li>표본이 {n_months}개월로 크지 않을 수 있어 참고 지표로만 활용하세요. 사업 구조가 크게 바뀐
        기업은 오래된 과거 평균이 왜곡 신호를 줄 수 있습니다.</li>
      <li>PER은 적자 구간(음수 EPS)에서 의미가 없어 해당 월은 자동 제외됩니다.</li>
      <li>"매수 타이밍 신호"(다이버전스)는 위 밸류에이션 판단과 완전히 별개입니다. 밸류에이션이
        "매우고평가"여도 반도체 사이클처럼 이익 개선이 기대되는 국면에서는 기술적으로 매수세 전환
        신호가 먼저 나타날 수 있습니다. 반대로 밸류에이션이 싸도 다이버전스 없이 계속 하락할 수도
        있습니다.</li>
      <li>다이버전스는 확정적 매매 신호가 아니라 패턴 기반 참고 지표입니다. 실제 매수 시에는 분할매수,
        손절 기준(예: 직전 저점 이탈 시 손절) 등 별도의 리스크 관리 규칙과 함께 사용하세요.</li>
      <li>"애널리스트 목표주가 &amp; 선행 밸류에이션"은 KIS API가 아니라 사용자가 증권사 리포트를 보고
        직접 입력한 값입니다. 리포트 발행 시점 기준 데이터라 시간이 지나면 낡을 수 있으니, 표시된
        증권사·시점을 꼭 함께 확인하세요. 목표주가는 해당 증권사의 전망치일 뿐 실현을 보장하지 않습니다.</li>
      <li>"과거 급등 이벤트 백테스트"의 이벤트는 실제 뉴스/실적 캘린더가 아니라 "하루 등락률+거래량이
        동시에 크게 튄 날"을 호재의 대리 지표로 삼아 자동 탐지한 것입니다. 표본 수가 적을 수 있고,
        과거 패턴이 미래에 반복된다는 보장도 없습니다. 참고 지표로만 활용하세요.</li>
      <li>"PER-주가 상관관계"는 PER=주가÷EPS라는 정의상 같은 회계연도 구간 안에서는 기계적으로
        연동되는 성격이 있어, 상관계수 자체보다 "PER 구간별 이후 수익률"(사후편향 없이 과거 시점
        기준으로 계산)이 더 실질적인 정보입니다. 표본이 적은 종목은 구간별 결과가 들쭉날쭉할 수
        있습니다.</li>
      <li>"외국인·프로그램매매 수급"의 당일 수치는 KIS 실시간 스냅샷이라 신뢰할 수 있지만, 최근
        추이 조회는 상대적으로 검증이 덜 된 API라 조회에 실패하면 그 부분만 조용히 생략될 수
        있습니다.</li>
    </ul>
  </div>

  <footer>개미는 뚠뚠 · {snapshot_date} 생성</footer>
</div>
</body>
</html>
"""


def fmt_x(v):
    return f"{v:.1f}x" if v is not None else "N/A"


def fmt_pct(v):
    return f"{v*100:.0f}%" if v is not None else "N/A"


def fmt_gap(cur, avg):
    if cur is None or avg is None or avg == 0:
        return "N/A"
    gap = (cur - avg) / avg
    sign = "+" if gap >= 0 else ""
    return f"{sign}{gap*100:.1f}%"


def fmt_r(r):
    return f"{r:.2f}" if r is not None else "-"


def fmt_signed_pct(v):
    if v is None:
        return "N/A"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v*100:.1f}%"


# --------------------------------------------------------------------- 외국인·프로그램매매 수급
# 당일 스냅샷(외국인 보유율/순매수, 프로그램매매 순매수)은 현재가 조회(quote) 응답에 이미 들어있는
# 필드라 항상 신뢰할 수 있다. 최근 며칠간의 추이는 별도 엔드포인트(get_investor_trend)를 추가로
# 호출해야 하는데, 이 엔드포인트는 다른 것들만큼 실전 검증을 못 해봤다 - 응답 스키마가 다르면
# 그 부분만 조용히 생략되고(예외로 잡힘) 당일 스냅샷은 정상 표시된다.

def render_investor_flow_section(quote, investor_trend_body=None):
    def to_float(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    frgn_ehrt_v = to_float(quote.get("hts_frgn_ehrt"))
    frgn_ntby_v = to_float(quote.get("frgn_ntby_qty"))
    pgtr_ntby_v = to_float(quote.get("pgtr_ntby_qty"))

    if frgn_ehrt_v is None and frgn_ntby_v is None:
        return ""  # quote에 관련 필드 자체가 없으면 섹션 생략

    def qty_badge(v):
        if v is None:
            return "N/A", "#898781"
        if v > 0:
            return f"+{v:,.0f}주 (순매수)", "#0ca30c"
        if v < 0:
            return f"{v:,.0f}주 (순매도)", "#d03b3b"
        return "0주 (보합)", "#898781"

    frgn_text, frgn_color = qty_badge(frgn_ntby_v)
    pgtr_text, pgtr_color = qty_badge(pgtr_ntby_v)

    snapshot_html = f"""
    <div class="stat-grid" style="grid-template-columns: repeat(3,1fr);">
      <div class="stat-tile"><div class="label">외국인 보유율</div><div class="value">{f"{frgn_ehrt_v:.2f}%" if frgn_ehrt_v is not None else 'N/A'}</div></div>
      <div class="stat-tile"><div class="label">당일 외국인 순매수</div><div class="value" style="font-size:15px;color:{frgn_color}">{frgn_text}</div></div>
      <div class="stat-tile"><div class="label">당일 프로그램매매 순매수</div><div class="value" style="font-size:15px;color:{pgtr_color}">{pgtr_text}</div></div>
    </div>"""

    trend_html = ""
    try:
        rows = investor_trend_body.get("output", []) if investor_trend_body else []
        parsed = []
        for r in rows:
            d = r.get("stck_bsop_date")
            f = r.get("frgn_ntby_qty")
            if d and f not in (None, ""):
                parsed.append((d, float(f)))
        if len(parsed) >= 2:
            parsed.sort()
            recent = parsed[-20:]
            n_days = len(recent)
            n_buy_days = sum(1 for _, v in recent if v > 0)
            cum = sum(v for _, v in recent)
            cum_text, cum_color = qty_badge(cum)
            max_abs = max(abs(v) for _, v in recent) or 1
            bar_rows = []
            for d, v in recent[-10:]:
                w = abs(v) / max_abs * 100
                color = "#0ca30c" if v > 0 else "#d03b3b" if v < 0 else "#898781"
                bar_rows.append(
                    f'<div style="display:flex;align-items:center;gap:8px;font-size:11px;color:var(--muted);">'
                    f'<span style="width:52px;">{d[4:6]}/{d[6:8]}</span>'
                    f'<div style="flex:1;background:var(--page);border-radius:3px;height:10px;">'
                    f'<div style="width:{w:.0f}%;background:{color};height:10px;border-radius:3px;"></div></div>'
                    f'<span style="width:76px;text-align:right;color:{color};">{v:+,.0f}주</span></div>'
                )
            trend_html = f"""
    <div class="tech-subtitle">최근 {n_days}거래일 외국인 순매수 추이 (그중 순매수일 {n_buy_days}일)</div>
    <div class="corr-range">누적 순매수: <span style="color:{cum_color};font-weight:700;">{cum_text}</span></div>
    <div style="display:flex;flex-direction:column;gap:4px;margin-top:8px;">
      {''.join(bar_rows)}
    </div>"""
    except Exception:
        trend_html = ""

    if not trend_html:
        trend_html = ('<div class="muted" style="margin-top:10px;">최근 며칠간 추이 데이터는 조회하지 못했습니다 '
                        '(오늘자 스냅샷만 표시 - 리포트의 다른 부분에는 영향 없습니다).</div>')

    return f"""
  <div class="card">
    <h2>외국인·프로그램매매 수급</h2>
    {snapshot_html}
    {trend_html}
  </div>"""


# --------------------------------------------------------------------- 애널리스트 목표주가 (수동 입력)
# analyst_targets.json (같은 폴더)에 사용자가 증권사 리포트를 보고 직접 입력해두는 데이터.
# KIS API로 자동 조회되는 값이 아니다 - 새 리포트가 나오면 이 파일을 직접 갱신해야 반영된다.
# customs_<label>.json과 같은 패턴: 파일이 없거나 해당 종목 항목이 없으면 섹션 자체가 생략된다.

def load_analyst_targets(code):
    fname = "analyst_targets.json"
    if not os.path.exists(fname):
        return []
    try:
        with open(fname, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return []
    return d.get(code, [])


def render_analyst_target_section(code, current_price):
    entries = load_analyst_targets(code)
    if not entries:
        return ""

    blocks = []
    for e in entries:
        broker = e.get("broker", "-")
        etype = e.get("type")

        if etype == "target_price":
            target_price = e.get("target_price")
            metric = e.get("metric", "")
            target_year = e.get("target_year", "")
            target_multiple = e.get("target_multiple")
            note = e.get("note", "")
            upside = (target_price - current_price) / current_price if target_price and current_price else None
            if upside is None:
                badge_text, badge_color = "N/A", "#898781"
            else:
                badge_text = f"{'업사이드' if upside >= 0 else '다운사이드'} {fmt_pct(abs(upside))}"
                badge_color = "#0ca30c" if upside >= 0 else "#d03b3b"
            blocks.append(f"""
    <div class="corr-product">
      <h3>{broker} <span class="corr-badge" style="background:{badge_color}">{badge_text}</span></h3>
      <div class="corr-range">목표주가 {target_price:,.0f}원 ({target_year}년 {metric} {target_multiple}배 기준) · 현재가 {current_price:,.0f}원 대비</div>
      {f'<div class="tech-explain">{note}</div>' if note else ''}
    </div>""")

        elif etype == "fwd_multiple_band":
            metric = e.get("metric", "")
            horizon = e.get("horizon", "")
            cur_v = e.get("current_value")
            avg_v = e.get("historical_avg")
            band_note = e.get("band_note", "")
            gap = (cur_v - avg_v) / avg_v if (cur_v is not None and avg_v) else None
            if gap is None:
                badge_text, badge_color = "N/A", "#898781"
            elif gap > 0:
                badge_text, badge_color = "평균 대비 고평가", "#d03b3b"
            else:
                badge_text, badge_color = "평균 대비 저평가", "#0ca30c"
            blocks.append(f"""
    <div class="corr-product">
      <h3>{broker} <span class="corr-badge" style="background:{badge_color}">{badge_text}</span></h3>
      <div class="corr-range">{horizon} {metric} {cur_v}배 vs 과거평균 {avg_v}배 ({fmt_gap(cur_v, avg_v)}){' · ' + band_note if band_note else ''}</div>
    </div>""")

    if not blocks:
        return ""

    return f"""
  <div class="card">
    <h2>애널리스트 목표주가 &amp; 선행 밸류에이션 (수동 입력)</h2>
    <div class="muted" style="margin-bottom:14px;">증권사 리포트를 보고 사용자가 직접 입력한 데이터입니다
    (analyst_targets.json 파일을 수정하면 갱신됩니다). KIS API로 실시간 조회되는 값이 아니라 입력 시점
    기준이니 날짜를 확인하세요. 업사이드는 위 목표주가와 KIS 실시간 현재가를 비교해 자동 계산됩니다.</div>
    {''.join(blocks)}
  </div>"""


def render_export_corr_section(code, bars):
    """STOCK_TO_PRODUCTS에 있는 종목이고 관련 customs_<label>.json이 로컬에 있으면
    수출입데이터 상관관계 카드를 렌더링. 없으면 빈 문자열(섹션 자체를 생략)."""
    products = STOCK_TO_PRODUCTS.get(code, [])
    if not products:
        return ""

    blocks = []
    for label_kr, label in products:
        cdict = load_customs_local(label)
        if cdict is None:
            print(f"  [안내] customs_{label}.json 이 없어 '{label_kr}' 상관관계 섹션은 생략합니다 "
                  f"(customs_export.py로 먼저 받아두면 표시됩니다).")
            continue
        corr = export_corr_analysis(bars, cdict)
        if corr is None:
            print(f"  [안내] '{label_kr}' 데이터와 주가의 공통 구간이 너무 짧아 상관관계를 계산하지 못했습니다.")
            continue

        rows_html = []
        for metric in ["수출액", "물량", "단가"]:
            m = corr[metric]
            rows_html.append(
                f"<tr><td>{metric}</td><td>{fmt_r(m['전체기간'])}</td>"
                f"<td>{fmt_r(m['2024-01~'])}</td><td>{fmt_r(m['2025-01~'])}</td></tr>"
            )
        # 헤드라인: 2025~ 구간에서 절대값이 가장 큰 지표를 대표로 표시
        best_metric, best_r = None, None
        for metric in ["수출액", "물량", "단가"]:
            r = corr[metric]["2025-01~"] or corr[metric]["전체기간"]
            if r is not None and (best_r is None or abs(r) > abs(best_r)):
                best_metric, best_r = metric, r
        vtext, vcolor = corr_verdict(best_r)

        blocks.append(f"""
    <div class="corr-product">
      <h3>{label_kr} <span class="corr-badge" style="background:{vcolor}">{vtext}</span></h3>
      <div class="corr-range">공통 구간 {corr['n']}개월 ({corr['range'][0]} ~ {corr['range'][1]}) · 대표 지표: {best_metric or '-'} (r={fmt_r(best_r)})</div>
      <table class="corr-table">
        <tr><th>지표</th><th>전체기간</th><th>2024~</th><th>2025~</th></tr>
        {''.join(rows_html)}
      </table>
    </div>""")

    if not blocks:
        return ""

    return f"""
  <div class="card">
    <h2>수출입데이터 상관관계 (관세청 품목별 수출입실적)</h2>
    <div class="muted" style="margin-bottom:14px;">상관계수(r)는 주가와 각 지표 간 Pearson 상관계수. |r|≥0.5 의미있는 신호,
    0.3~0.5 약한 신호, 그 미만 무의미로 표시합니다. DRAM/NAND처럼 소수 기업이 수출을 사실상 독점하는
    품목은 신호가 강하고, MLCC/카메라모듈처럼 여러 회사가 섞인 품목은 신호가 약할 수 있습니다.</div>
    {''.join(blocks)}
  </div>"""


def render_technical_signal_section(daily_bars):
    """daily_bars: get_daily_bars_chunked() 반환값 (날짜 오름차순 [{"date","close","volume"}, ...]).
    데이터가 없거나 너무 적으면 빈 문자열(섹션 생략)."""
    MIN_BARS = 40
    if not daily_bars or len(daily_bars) < MIN_BARS:
        return ""

    dates = [b["date"] for b in daily_bars]
    closes = [b["close"] for b in daily_bars]
    volumes = [b["volume"] for b in daily_bars]
    obv = compute_obv(closes, volumes)
    rsi = compute_rsi(closes, period=14)

    result = detect_divergence(dates, closes, obv, rsi, order=5)
    title, color, explain = divergence_verdict(result)

    def fmt_date(d):
        return f"{d[4:6]}/{d[6:8]}"

    cur_price = closes[-1]
    ma20 = moving_average(closes, 20)
    ma60 = moving_average(closes, 60)
    cur_rsi = next((v for v in reversed(rsi) if v is not None), None)
    rsi_zone = "-"
    if cur_rsi is not None:
        rsi_zone = "과매수(70+)" if cur_rsi >= 70 else "과매도(30-)" if cur_rsi <= 30 else "중립"

    def ma_badge(price, ma):
        if ma is None:
            return "N/A"
        return f"{ma:,.0f}원 ({'현재가 상회' if price >= ma else '현재가 하회'})"

    charts_html = ""
    table_html = ""
    if result["points"]:
        p = result["points"]
        i1 = dates.index(p["date1"])
        i2 = dates.index(p["date2"])
        price_chart = swing_chart_svg(dates, closes, [i1, i2],
                                       lambda d, v: f"{fmt_date(d)} {v:,.0f}원")
        obv_chart = swing_chart_svg(dates, obv, [i1, i2],
                                     lambda d, v: f"{fmt_date(d)} {v:,.0f}")
        point_label = "저점" if result["type"] == "bullish" else "고점"
        charts_html = f"""
    <div class="tech-subtitle">주가 추이 (직전 {point_label} <span style="color:#898781">●</span> vs 최근 {point_label} <span style="color:#2a78d6">●</span>)</div>
    {price_chart}
    <div class="tech-subtitle">OBV(거래량 누적지표) 추이 — 같은 두 시점 비교</div>
    {obv_chart}
    <table class="corr-table" style="margin-top:10px;">
      <tr><th>시점</th><th>날짜</th><th>종가</th><th>OBV</th><th>RSI(14)</th></tr>
      <tr><td>직전 {point_label}</td><td>{p['date1']}</td><td>{p['price1']:,.0f}원</td><td>{p['obv1']:,.0f}</td><td>{f"{p['rsi1']:.0f}" if p['rsi1'] is not None else '-'}</td></tr>
      <tr><td>최근 {point_label}</td><td>{p['date2']}</td><td>{p['price2']:,.0f}원</td><td>{p['obv2']:,.0f}</td><td>{f"{p['rsi2']:.0f}" if p['rsi2'] is not None else '-'}</td></tr>
    </table>"""

    return f"""
  <div class="card">
    <h2>매수 타이밍 신호 (기술적 분석 — 다이버전스)</h2>
    <div class="verdict-row">
      <span class="corr-badge" style="background:{color}; font-size:14px; padding:6px 14px;">{title}</span>
    </div>
    <div class="tech-explain">{explain}</div>
    <div class="stat-grid" style="grid-template-columns: repeat(4,1fr);">
      <div class="stat-tile"><div class="label">현재가</div><div class="value">{cur_price:,.0f}원</div></div>
      <div class="stat-tile"><div class="label">20일 이동평균</div><div class="value" style="font-size:14px;">{ma_badge(cur_price, ma20)}</div></div>
      <div class="stat-tile"><div class="label">60일 이동평균</div><div class="value" style="font-size:14px;">{ma_badge(cur_price, ma60)}</div></div>
      <div class="stat-tile"><div class="label">RSI(14)</div><div class="value">{f"{cur_rsi:.0f}" if cur_rsi is not None else 'N/A'} <span style="font-size:11px;color:var(--muted);">{rsi_zone}</span></div></div>
    </div>
    {charts_html}
    <div class="tech-legend">일봉 {len(daily_bars)}개 ({dates[0]} ~ {dates[-1]}) 기준 · 이 신호는 위 밸류에이션(PER/PBR) 판단과 별개입니다.</div>
  </div>"""


# --------------------------------------------------------------------- 과거 급등 이벤트 백테스트
# event_backtest.py를 미리 실행해서 만들어둔 event_backtest_<code>.json (같은 폴더)이 있을 때만 표시.
# "하루 등락률+거래량이 동시에 크게 튄 날"을 호재/실적서프라이즈의 대리 지표로 삼아, 그 이후
# 수익률 궤적과 사전 다이버전스 여부를 과거 여러 건에 대해 집계한 결과 - 실제 뉴스 캘린더 기반이
# 아니므로 event_history.json에 확인된 사건이 있으면 라벨만 참고용으로 붙는다.

def load_event_backtest(code):
    fname = f"event_backtest_{code}.json"
    if not os.path.exists(fname):
        return None
    try:
        with open(fname, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def render_event_backtest_section(code, daily_bars):
    result = load_event_backtest(code)
    if result is None or not result.get("n_events"):
        return ""

    cur_state_html = "라이브 일봉 데이터가 없어 지금 상태와는 비교할 수 없습니다."
    if daily_bars and len(daily_bars) >= 40:
        dates = [b["date"] for b in daily_bars]
        closes = [b["close"] for b in daily_bars]
        vols = [b["volume"] for b in daily_bars]
        obv = compute_obv(closes, vols)
        rsi = compute_rsi(closes)
        cur_result = detect_divergence(dates, closes, obv, rsi, order=5)
        if cur_result["type"] == "bullish":
            n_confirm = sum(1 for s in cur_result["signals"].values() if s)
            cur_state_html = (f"지금도 강세 다이버전스 신호가 {'뚜렷하게' if n_confirm == 2 else '약하게'} "
                               f"포착된 상태입니다 (위 '매수 타이밍 신호' 섹션과 같은 로직).")
        else:
            cur_state_html = "지금은 다이버전스 신호가 뚜렷하지 않은 상태입니다."

    agg_rows = []
    for h, label in [("5", "T+5일"), ("10", "T+10일"), ("20", "T+20일"), ("40", "T+40일")]:
        a = result["aggregate"].get(h, {})
        win = fmt_pct(a.get("win_rate")) if a.get("win_rate") is not None else "-"
        agg_rows.append(
            f"<tr><td>{label}</td><td>{a.get('n', '-')}</td>"
            f"<td>{fmt_signed_pct(a.get('mean'))}</td><td>{fmt_signed_pct(a.get('median'))}</td><td>{win}</td></tr>"
        )

    event_rows = []
    for e in result["events"][-10:]:
        label = (f' <span class="corr-badge" style="background:#2a78d6;">{e["label"]}</span>'
                  if e.get("label") else "")
        pre = "O" if e.get("pre_divergence_bullish") else "-"
        event_rows.append(
            f"<tr><td>{e['date']}{label}</td><td>{fmt_signed_pct(e['ret'])}</td>"
            f"<td>{e['vol_ratio']:.1f}배</td><td>{pre}</td><td>{fmt_signed_pct(e['forward'].get('20'))}</td></tr>"
        )

    return f"""
  <div class="card">
    <h2>과거 급등 이벤트 백테스트 (호재/실적서프라이즈 대리 지표)</h2>
    <div class="muted" style="margin-bottom:14px;">하루 {result['min_abs_return']*100:.0f}%+ 상승 &amp;
    거래량 평균 대비 {result['min_vol_ratio']:.1f}배+ 인 날을 과거 데이터에서 자동으로 찾아 그 이후
    수익률을 집계했습니다 (event_backtest.py로 갱신). 실제 뉴스 캘린더가 아니라 가격·거래량 패턴 기반
    대리 지표이니 참고용으로만 활용하세요. 데이터 구간: {result['data_range'][0]} ~ {result['data_range'][1]}
    · 탐지된 이벤트 {result['n_events']}건 (사전 강세다이버전스 동반: {result['n_pre_divergence_bullish']}건)</div>
    <table class="corr-table">
      <tr><th>기간</th><th>표본</th><th>평균수익률</th><th>중앙값</th><th>승률</th></tr>
      {''.join(agg_rows)}
    </table>
    <div class="tech-subtitle">최근 이벤트 (최대 10건, 파란 라벨은 event_history.json에 등록된 실제 뉴스와 날짜가 겹치는 경우)</div>
    <table class="corr-table">
      <tr><th>날짜</th><th>당일수익률</th><th>거래량</th><th>사전다이버전스</th><th>T+20수익률</th></tr>
      {''.join(event_rows)}
    </table>
    <div class="tech-explain" style="margin-top:12px;">{cur_state_html}</div>
  </div>"""


def build_report(code, name, quote, annual, bars, daily_bars=None, investor_trend=None):
    history = build_history(annual, bars)
    n_months = len(history)
    if n_months < 6:
        sys.exit(f"과거 데이터가 {n_months}개월뿐이라 리포트를 만들기 어렵습니다 (최소 6개월 필요). "
                  f"최근 상장했거나 재무비율 데이터가 부족한 종목일 수 있습니다.")

    per_series = [row[2] for row in history]
    pbr_series = [row[3] for row in history]
    cur_per = float(quote["per"]) if quote.get("per") not in (None, "") else None
    cur_pbr = float(quote["pbr"]) if quote.get("pbr") not in (None, "") else None
    price = float(quote["stck_prpr"])

    avg_per = statistics.mean([v for v in per_series if v is not None]) if any(v is not None for v in per_series) else None
    avg_pbr = statistics.mean([v for v in pbr_series if v is not None]) if any(v is not None for v in pbr_series) else None

    per_pct = percentile_of(cur_per, per_series)
    pbr_pct = percentile_of(cur_pbr, pbr_series)
    valid_pcts = [p for p in (per_pct, pbr_pct) if p is not None]
    composite_pct = statistics.mean(valid_pcts) if valid_pcts else None
    verdict = verdict_of(composite_pct)
    verdict_color = VERDICT_COLOR[verdict][0]

    per_gauge = gauge_svg(per_pct, verdict_of(per_pct))
    pbr_gauge = gauge_svg(pbr_pct, verdict_of(pbr_pct))
    per_chart = f'<div class="gauge-title">PER 추이 (과거 {n_months}개월)</div>' + line_chart_svg(history, 2, "PER")
    pbr_chart = f'<div class="gauge-title">PBR 추이 (과거 {n_months}개월)</div>' + line_chart_svg(history, 3, "PBR")
    export_corr_section = render_export_corr_section(code, bars)
    technical_section = render_technical_signal_section(daily_bars)
    analyst_section = render_analyst_target_section(code, price)
    event_backtest_section = render_event_backtest_section(code, daily_bars)
    per_price_corr_section = render_per_price_correlation_section(history)
    investor_flow_section = render_investor_flow_section(quote, investor_trend)

    from datetime import datetime
    snapshot_date = datetime.now().strftime("%Y-%m-%d %H:%M")

    html = HTML_TEMPLATE.format(
        name=name, code=code, snapshot_date=snapshot_date,
        verdict=verdict, verdict_color=verdict_color, composite_pct=fmt_pct(composite_pct),
        price=price, cur_per=fmt_x(cur_per), cur_pbr=fmt_x(cur_pbr), n_months=n_months,
        avg_per=fmt_x(avg_per), avg_pbr=fmt_x(avg_pbr),
        per_gap=fmt_gap(cur_per, avg_per), pbr_gap=fmt_gap(cur_pbr, avg_pbr),
        per_gauge=per_gauge, pbr_gauge=pbr_gauge, per_chart=per_chart, pbr_chart=pbr_chart,
        export_corr_section=export_corr_section, technical_section=technical_section,
        analyst_section=analyst_section, event_backtest_section=event_backtest_section,
        per_price_corr_section=per_price_corr_section, investor_flow_section=investor_flow_section,
    )
    return html, verdict, composite_pct


def main():
    if len(sys.argv) >= 2:
        code = sys.argv[1]
        name = sys.argv[2] if len(sys.argv) >= 3 else code
    else:
        code = input("종목코드(6자리)를 입력하세요: ").strip()
        name = input("종목명을 입력하세요 (생략 가능, 엔터): ").strip() or code

    appkey, appsecret = get_credentials()
    print("Requesting access token...")
    token = get_access_token(appkey, appsecret)
    print(f"[{code} {name}] 데이터 조회 중...")

    _, quote = _call_with_retry(get_current_quote, token, appkey, appsecret, code)
    if quote.get("rt_cd") != "0" or not quote.get("output"):
        sys.exit(f"현재가 조회 실패: {quote}")
    time.sleep(0.5)

    _, fr = _call_with_retry(get_financial_ratio, token, appkey, appsecret, code)
    rows = fr.get("output", [])
    annual = {int(r["stac_yymm"][:4]): (float(r["eps"]), float(r["bps"]))
              for r in rows if r["stac_yymm"][4:] == "12" and r.get("eps") and r.get("bps")}
    if not annual:
        sys.exit("연간 재무비율(EPS/BPS) 데이터를 찾을 수 없습니다. 종목코드를 확인해주세요.")
    time.sleep(0.5)

    from datetime import datetime as _dt
    end_str = _dt.now().strftime("%Y%m%d")
    _, ph = _call_with_retry(get_price_history, token, appkey, appsecret, code, "20190101", end_str, "M")
    bars = ph.get("output2", [])
    if len(bars) < 6:
        sys.exit(f"월별 주가 데이터가 {len(bars)}개월뿐입니다. 종목코드를 확인해주세요.")
    time.sleep(0.5)

    print("매수 타이밍 신호용 일봉 데이터 조회 중... (여러 번 호출해서 시간이 조금 걸립니다)")
    try:
        daily_bars = get_daily_bars_chunked(token, appkey, appsecret, code, lookback_days=400)
    except Exception as e:
        print(f"  [안내] 일봉 데이터 조회에 실패해 '매수 타이밍 신호' 섹션은 생략합니다: {e}")
        daily_bars = None

    try:
        _, investor_trend = _call_with_retry(get_investor_trend, token, appkey, appsecret, code)
    except Exception:
        investor_trend = None

    html, verdict, composite_pct = build_report(code, name, quote["output"], annual, bars, daily_bars, investor_trend)

    out_filename = f"개미는뚠뚠_리포트_{code}.html"
    with open(out_filename, "w", encoding="utf-8") as f:
        f.write(html)

    pct_str = f"{composite_pct*100:.0f}%ile" if composite_pct is not None else "N/A"
    print(f"\n=== {name}({code}) : {verdict} (종합 백분위 {pct_str}) ===")
    print(f"리포트 저장됨: {out_filename}  (더블클릭해서 브라우저로 열어보세요)")


if __name__ == "__main__":
    main()
