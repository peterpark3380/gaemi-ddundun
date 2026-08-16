"""
개미는 뚠뚠 - 웹앱 버전
================================================================
gaemi_ddundun.py (로컬 CLI 버전)의 핵심 로직을 그대로 재사용해서, 폰 브라우저에서
종목코드를 검색하면 리포트가 바로 뜨는 웹앱으로 만든 버전.

로컬에서 테스트:
  pip install -r requirements.txt
  set KIS_APPKEY=...
  set KIS_APPSECRET=...
  python3 app.py
  브라우저에서 http://localhost:5000 접속

배포:
  README_배포방법.md 참고 (Render.com 무료 호스팅 기준)
  KIS_APPKEY / KIS_APPSECRET 은 절대 코드에 넣지 말고, 호스팅 서비스의
  환경변수(Environment Variables) 설정 화면에만 입력할 것.
"""

import os
from datetime import datetime
from flask import Flask, request, Response

import gaemi_ddundun as core

app = Flask(__name__)

SEARCH_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>개미는 뚠뚠</title>
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/static/icon-192.png">
<meta name="theme-color" content="#0b0b0b">
<style>
  :root {{ color-scheme: light dark; }}
  * {{ box-sizing: border-box; }}
  body {{
    margin:0; font-family: system-ui,-apple-system,"Segoe UI",sans-serif;
    background: #f9f9f7; color:#0b0b0b;
    display:flex; align-items:center; justify-content:center; min-height:100vh; padding:20px;
  }}
  @media (prefers-color-scheme: dark) {{
    body {{ background:#0d0d0d; color:#fff; }}
    .card {{ background:#1a1a19 !important; border-color: rgba(255,255,255,0.10) !important; }}
    input {{ background:#0d0d0d !important; color:#fff !important; border-color: rgba(255,255,255,0.2) !important; }}
  }}
  .card {{ background:#fcfcfb; border:1px solid rgba(11,11,11,0.10); border-radius:16px; padding:32px 24px; max-width:380px; width:100%; text-align:center; }}
  .card h1 {{ font-size:26px; margin: 0 0 4px; }}
  .card p {{ color:#898781; font-size:13px; margin: 0 0 24px; }}
  form {{ display:flex; flex-direction:column; gap:10px; }}
  input {{ font-size:16px; padding:12px 14px; border-radius:10px; border:1px solid #e1e0d9; }}
  button {{ font-size:16px; font-weight:700; padding:12px 14px; border-radius:10px; border:none; background:#2a78d6; color:#fff; cursor:pointer; }}
  button:active {{ opacity:0.85; }}
  .error {{ color:#d03b3b; font-size:13px; margin-top:14px; }}
  .examples {{ margin-top:18px; font-size:11px; color:#898781; line-height:1.6; }}
</style>
</head>
<body>
  <div class="card">
    <h1>🐜 개미는 뚠뚠</h1>
    <p>종목코드를 검색하면 과거 대비 싼지 비싼지 알려드려요</p>
    <form action="/report" method="get">
      <input name="code" placeholder="종목코드 6자리 (예: 005930)" required pattern="[0-9]{{6}}" maxlength="6" inputmode="numeric">
      <input name="name" placeholder="종목명 (선택, 예: 삼성전자)">
      <button type="submit">검색</button>
    </form>
    {error_html}
    <div class="examples">005930 삼성전자 · 000660 SK하이닉스 · 009150 삼성전기 · 011070 LG이노텍</div>
  </div>
</body>
</html>"""


@app.route("/")
def index():
    return SEARCH_PAGE.format(error_html="")


@app.route("/manifest.json")
def manifest():
    return {
        "name": "개미는 뚠뚠",
        "short_name": "개미는뚠뚠",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#f9f9f7",
        "theme_color": "#0b0b0b",
        "icons": [
            {"src": "/static/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "/static/icon-512.png", "sizes": "512x512", "type": "image/png"},
        ],
    }


@app.route("/report")
def report():
    code = request.args.get("code", "").strip()
    name = request.args.get("name", "").strip() or code

    if not code.isdigit() or len(code) != 6:
        return SEARCH_PAGE.format(error_html='<div class="error">6자리 종목코드를 입력해주세요.</div>')

    try:
        appkey, appsecret = os.environ.get("KIS_APPKEY"), os.environ.get("KIS_APPSECRET")
        if not appkey or not appsecret:
            return SEARCH_PAGE.format(
                error_html='<div class="error">서버에 KIS_APPKEY/KIS_APPSECRET이 설정되지 않았습니다. '
                           '호스팅 서비스의 환경변수 설정을 확인해주세요.</div>'
            )
        token = core.get_access_token(appkey, appsecret)

        _, quote = core._call_with_retry(core.get_current_quote, token, appkey, appsecret, code)
        if quote.get("rt_cd") != "0" or not quote.get("output"):
            msg = quote.get("msg1", "알 수 없는 오류")
            return SEARCH_PAGE.format(error_html=f'<div class="error">조회 실패: {msg} (종목코드를 확인해주세요)</div>')

        _, fr = core._call_with_retry(core.get_financial_ratio, token, appkey, appsecret, code)
        rows = fr.get("output", [])
        annual = {int(r["stac_yymm"][:4]): (float(r["eps"]), float(r["bps"]))
                  for r in rows if r["stac_yymm"][4:] == "12" and r.get("eps") and r.get("bps")}
        if not annual:
            return SEARCH_PAGE.format(error_html='<div class="error">연간 재무비율 데이터를 찾을 수 없습니다.</div>')

        end_date = datetime.now().strftime("%Y%m%d")
        _, ph = core._call_with_retry(core.get_price_history, token, appkey, appsecret, code, "20190101", end_date, "M")
        bars = ph.get("output2", [])
        if len(bars) < 6:
            return SEARCH_PAGE.format(error_html='<div class="error">월별 주가 데이터가 부족합니다.</div>')

        try:
            daily_bars = core.get_daily_bars_chunked(token, appkey, appsecret, code, lookback_days=400)
        except Exception:
            daily_bars = None

        try:
            _, investor_trend = core._call_with_retry(core.get_investor_trend, token, appkey, appsecret, code)
        except Exception:
            investor_trend = None

        html, verdict, composite_pct = core.build_report(code, name, quote["output"], annual, bars, daily_bars, investor_trend)
        return Response(html, mimetype="text/html")

    except SystemExit as e:
        return SEARCH_PAGE.format(error_html=f'<div class="error">{e}</div>')
    except Exception as e:
        return SEARCH_PAGE.format(error_html=f'<div class="error">오류가 발생했습니다: {e}</div>')


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
