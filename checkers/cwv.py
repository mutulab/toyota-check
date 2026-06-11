"""Core Web Vitals チェッカー（PageSpeed Insights API）"""

import json
import time
import urllib.request
import urllib.parse
from config import PSI_API_KEY, PSI_STRATEGY


PSI_ENDPOINT = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"

# TQP KPI 閾値
THRESHOLDS = {
    "LCP":  {"good": 2500,  "poor": 4000,  "unit": "ms"},
    "CLS":  {"good": 0.1,   "poor": 0.25,  "unit": ""},
    "INP":  {"good": 200,   "poor": 500,   "unit": "ms"},
    "FCP":  {"good": 1800,  "poor": 3000,  "unit": "ms"},
    "TTFB": {"good": 800,   "poor": 1800,  "unit": "ms"},
}


def _rate(value, metric: str) -> str:
    t = THRESHOLDS.get(metric, {})
    if not t or value is None:
        return "N/A"
    if value <= t["good"]:
        return "GOOD"
    elif value <= t["poor"]:
        return "NEEDS_IMPROVEMENT"
    return "POOR"


def check_cwv(url: str, strategy: str = PSI_STRATEGY) -> dict:
    """1URLのCWVを取得"""
    params = {"url": url, "strategy": strategy, "category": "performance"}
    if PSI_API_KEY:
        params["key"] = PSI_API_KEY
    query = urllib.parse.urlencode(params)
    api_url = f"{PSI_ENDPOINT}?{query}"

    try:
        req = urllib.request.Request(
            api_url,
            headers={"User-Agent": "toyota-check/1.0"},
        )
        with urllib.request.urlopen(req, timeout=30) as res:
            data = json.loads(res.read().decode())
    except Exception as e:
        return {"url": url, "error": str(e)}

    # Lighthouse メトリクス抽出
    cats = data.get("lighthouseResult", {})
    audits = cats.get("audits", {})
    score = cats.get("categories", {}).get("performance", {}).get("score")

    def ms(key):
        v = audits.get(key, {}).get("numericValue")
        return round(v) if v is not None else None

    def cls_val():
        v = audits.get("cumulative-layout-shift", {}).get("numericValue")
        return round(v, 3) if v is not None else None

    lcp  = ms("largest-contentful-paint")
    cls  = cls_val()
    inp  = ms("interaction-to-next-paint") or ms("total-blocking-time")
    fcp  = ms("first-contentful-paint")
    ttfb = ms("server-response-time")

    result = {
        "url": url,
        "strategy": strategy,
        "perf_score": round(score * 100) if score else None,
        "LCP":  lcp,  "LCP_rate":  _rate(lcp, "LCP"),
        "CLS":  cls,  "CLS_rate":  _rate(cls, "CLS"),
        "INP":  inp,  "INP_rate":  _rate(inp, "INP"),
        "FCP":  fcp,  "FCP_rate":  _rate(fcp, "FCP"),
        "TTFB": ttfb, "TTFB_rate": _rate(ttfb, "TTFB"),
        "ng_items": [],
    }
    for m in ["LCP", "CLS", "INP"]:
        if result[f"{m}_rate"] in ("NEEDS_IMPROVEMENT", "POOR"):
            result["ng_items"].append(m)

    return result


def check_cwv_bulk(urls: list[str], strategy: str = PSI_STRATEGY) -> list[dict]:
    """複数URLのCWVを順次取得（PSI APIはRateLimit対策で逐次処理）"""
    results = []
    for i, url in enumerate(urls, 1):
        print(f"  [{i}/{len(urls)}] {url}")
        r = check_cwv(url, strategy)
        results.append(r)
        time.sleep(1.0)  # PSI API のレート制限対策
    return results
