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

HOW TO RUN
  같은 cmd 창에 KIS_APPKEY / KIS_APPSECRET 이 이미 set 되어 있으면:
    python3 gaemi_ddundun.py 011070 LG이노텍
    python3 gaemi_ddundun.py 011070          (종목명 생략 가능)
    python3 gaemi_ddundun.py                 (인자 없이 실행하면 직접 입력받음)

  수출입데이터 상관관계까지 보려면 같은 폴더에 아래 파일들이 있어야 함
  (customs_export.py로 미리 받아둘 것):
    005930(삼성전자)/000660(SK하이닉스): customs_dram.json, customs_hbm.json, customs_nand.json
    009150(삼성전기): customs_mlcc.json, customs_camera.json
    011070(LG이노텍): customs_camera.json, customs_pcb_c.json
  파일이 없으면 그 섹션은 자동으로 생략되고 나머지 리포트는 정상 생성된다.

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
    --muted: #898781; --grid: #e1e0d9; --series-1: #2a78d6; --border: rgba(11,11,11,0.10);
  }}
  @media (prefers-color-scheme: dark) {{
    :root:where(:not([data-theme="light"])) .viz-root {{
      color-scheme: dark;
      --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff; --text-secondary: #c3c2b7;
      --muted: #898781; --grid: #2c2c2a; --series-1: #3987e5; --border: rgba(255,255,255,0.10);
    }}
  }}
  :root[data-theme="dark"] .viz-root {{
    color-scheme: dark;
    --surface-1: #1a1a19; --page: #0d0d0d; --text-primary: #ffffff; --text-secondary: #c3c2b7;
    --muted: #898781; --grid: #2c2c2a; --series-1: #3987e5; --border: rgba(255,255,255,0.10);
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
  .series-dot {{ fill: var(--series-1); stroke: var(--surface-1); stroke-width: 2; }}
  .cur-label {{ font-size: 11px; font-weight:700; fill: var(--series-1); }}
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


def build_report(code, name, quote, annual, bars):
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

    from datetime import datetime
    snapshot_date = datetime.now().strftime("%Y-%m-%d %H:%M")

    html = HTML_TEMPLATE.format(
        name=name, code=code, snapshot_date=snapshot_date,
        verdict=verdict, verdict_color=verdict_color, composite_pct=fmt_pct(composite_pct),
        price=price, cur_per=fmt_x(cur_per), cur_pbr=fmt_x(cur_pbr), n_months=n_months,
        avg_per=fmt_x(avg_per), avg_pbr=fmt_x(avg_pbr),
        per_gap=fmt_gap(cur_per, avg_per), pbr_gap=fmt_gap(cur_pbr, avg_pbr),
        per_gauge=per_gauge, pbr_gauge=pbr_gauge, per_chart=per_chart, pbr_chart=pbr_chart,
        export_corr_section=export_corr_section,
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

    _, ph = _call_with_retry(get_price_history, token, appkey, appsecret, code, "20190101", "20260814", "M")
    bars = ph.get("output2", [])
    if len(bars) < 6:
        sys.exit(f"월별 주가 데이터가 {len(bars)}개월뿐입니다. 종목코드를 확인해주세요.")

    html, verdict, composite_pct = build_report(code, name, quote["output"], annual, bars)

    out_filename = f"개미는뚠뚠_리포트_{code}.html"
    with open(out_filename, "w", encoding="utf-8") as f:
        f.write(html)

    pct_str = f"{composite_pct*100:.0f}%ile" if composite_pct is not None else "N/A"
    print(f"\n=== {name}({code}) : {verdict} (종합 백분위 {pct_str}) ===")
    print(f"리포트 저장됨: {out_filename}  (더블클릭해서 브라우저로 열어보세요)")


if __name__ == "__main__":
    main()
