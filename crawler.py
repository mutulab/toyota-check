"""バックグラウンドクロールジョブ管理"""
from __future__ import annotations

import json
import re
import sys
import time
import threading
import uuid
from collections import deque
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

_JOBS: dict = {}
_DIR = Path("/tmp/toyota-check-jobs")
_DIR.mkdir(parents=True, exist_ok=True)

# ── Public API ──────────────────────────────────────────────────────────────

def start_job(cfg: dict) -> str:
    """ジョブを起動してジョブIDを返す"""
    job_id = uuid.uuid4().hex[:8].upper()
    job = {
        "id": job_id,
        "status": "running",
        "phase": "開始待機",
        "progress": 0,
        "cfg": cfg,
        "started_at": datetime.now().isoformat(),
        "finished_at": None,
        "last_updated_at": datetime.now().isoformat(),
        "url_count": 0,
        "urls": [],
        "results": {},
        "error": None,
    }
    _JOBS[job_id] = job
    _save(job_id)
    threading.Thread(target=_run, args=(job_id,), daemon=False).start()
    return job_id


def get_job(job_id: str) -> dict | None:
    """ジョブIDでジョブを取得。メモリになければファイルから読む"""
    jid = job_id.strip().upper()
    if jid in _JOBS:
        return _JOBS[jid]
    p = _DIR / f"{jid}.json"
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            _JOBS[jid] = data
            return data
        except Exception:
            pass
    return None


def list_jobs() -> list[dict]:
    """最近のジョブ一覧（最大20件）"""
    jobs = []
    for p in sorted(_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True)[:20]:
        try:
            jobs.append(json.loads(p.read_text(encoding="utf-8")))
        except Exception:
            pass
    return jobs


# ── Internal ────────────────────────────────────────────────────────────────

def _save(job_id: str):
    try:
        (_DIR / f"{job_id}.json").write_text(
            json.dumps(_JOBS[job_id], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _upd(job_id: str, **kw):
    _JOBS[job_id].update(kw)
    _JOBS[job_id]["last_updated_at"] = datetime.now().isoformat()
    _save(job_id)


def _run(job_id: str):
    root = str(Path(__file__).parent)
    if root not in sys.path:
        sys.path.insert(0, root)

    from concurrent.futures import ThreadPoolExecutor, as_completed
    from fetcher import (fetch_html, extract_links, extract_resources,
                         extract_css_urls, extract_meta, extract_text_blocks)
    from config import (REQUEST_DELAY, THRESHOLDS, PSI_ENDPOINT, KINSHI_PATTERNS)
    from dict_loader import load_for_check

    job = _JOBS[job_id]
    cfg = job["cfg"]

    try:
        start_url   = cfg.get("start_url", "")
        max_pages   = cfg["max_pages"]
        max_depth   = cfg["depth"]
        check_types = set(cfg.get("check_types", []))
        toyota_only = cfg.get("toyota_only", True)
        sel_res     = set(cfg.get("selected_res", []))
        psi_key     = cfg.get("psi_key", "")
        strategy    = cfg.get("strategy", "mobile")
        custom_raw  = cfg.get("custom_dict", "")
        base_domain = urlparse(start_url).netloc
        # Excel mode: derive base_domain from first supplied URL
        if not base_domain and cfg.get("url_source_type") == "excel":
            _first = (cfg.get("urls") or [""])[0]
            base_domain = urlparse(_first).netloc

        # ── Phase 1: URL取得（クロール or Excel指定） ─────────────────────
        url_source_type = cfg.get("url_source_type", "crawl")

        if url_source_type == "excel":
            urls = list(cfg.get("urls", []))[:max_pages]
            _upd(job_id,
                 phase="URL読み込み完了（コンテンツ管理票）",
                 progress=30,
                 urls=urls[:],
                 url_count=len(urls))
        else:
            _upd(job_id, phase="クロール中", progress=0)
            visited: set = set()
            queue: deque = deque([(start_url.rstrip("/"), 0)])
            urls: list = []

            while queue and len(urls) < max_pages:
                url, depth = queue.popleft()
                norm = url.rstrip("/").split("?")[0].split("#")[0]
                if norm in visited:
                    continue
                visited.add(norm)
                code, html = fetch_html(url)
                if code == 200 and html:
                    urls.append(url)
                    _upd(job_id,
                         urls=urls[:],
                         url_count=len(urls),
                         progress=round(len(urls) / max_pages * 30))
                    if depth < max_depth:
                        for link in extract_links(html, url):
                            if urlparse(link).netloc == base_domain:
                                c = link.rstrip("/").split("?")[0].split("#")[0]
                                if c not in visited:
                                    queue.append((link, depth + 1))
                time.sleep(REQUEST_DELAY)

        n = len(urls)
        results: dict = {}
        workers = 3

        # ── Phase 2: リンクチェック ───────────────────────────────────────
        if "link" in check_types and n:
            _css_res = {"画像(CSS)", "フォント(CSS)", "CSSリソース"}

            def _collect(url):
                code, html = fetch_html(url)
                res = extract_resources(html, url) if (code == 200 and html) else []
                if not sel_res or sel_res & _css_res:
                    for cu in [u for u, t in res if t == "CSS/スタイル"]:
                        cc, ct = fetch_html(cu)
                        if cc == 200 and ct:
                            res.extend(extract_css_urls(ct, cu))
                if sel_res:
                    res = [(u, t) for u, t in res if t in sel_res]
                if toyota_only:
                    res = [(u, t) for u, t in res if urlparse(u).netloc == base_domain]
                return url, res

            page_res: dict = {}
            done = 0
            with ThreadPoolExecutor(max_workers=workers) as ex:
                futs = {ex.submit(_collect, u): u for u in urls}
                for f in as_completed(futs):
                    src, res = f.result()
                    page_res[src] = res
                    done += 1
                    _upd(job_id,
                         phase=f"リンク収集中 ({done}/{n})",
                         progress=30 + round(done / n * 20))
                    time.sleep(REQUEST_DELAY)

            all_ru = list({u for res in page_res.values() for u, _ in res})

            def _ping(url):
                c, _ = fetch_html(url)
                return url, c

            ls: dict = {}
            done = 0
            with ThreadPoolExecutor(max_workers=workers) as ex:
                futs = {ex.submit(_ping, u): u for u in all_ru}
                for f in as_completed(futs):
                    u, c = f.result()
                    ls[u] = c
                    done += 1
                    _upd(job_id,
                         phase=f"リンク確認中 ({done}/{len(all_ru)})",
                         progress=50 + round(done / max(len(all_ru), 1) * 15))
                    time.sleep(REQUEST_DELAY)

            broken = []
            for src in sorted(page_res):
                for ru, rt in page_res[src]:
                    c = ls.get(ru, 0)
                    if c in (404, 410) or c == 0:
                        broken.append({"発見ページ": src, "リソースURL": ru,
                                       "種別": rt, "ステータス": c})
            results["link"] = broken

        # ── Phase 3: 表記ゆれ ─────────────────────────────────────────────
        if "content" in check_types and n:
            merged = load_for_check()
            for line in custom_raw.splitlines():
                if "|" in line:
                    w, c = line.split("|", 1)
                    merged[re.escape(w.strip())] = c.strip()

            def _content(url):
                code, html = fetch_html(url)
                if code != 200 or not html:
                    return []
                meta = extract_meta(html)
                text = " ".join(extract_text_blocks(html))
                rows = []
                for pat, rec in merged.items():
                    for m in sorted(set(re.findall(pat, text, re.I))):
                        rows.append({"URL": url, "ページタイトル": meta["short_title"],
                                     "種別": "表記ゆれ", "発見テキスト": m,
                                     "推奨表記": rec, "修正済み": ""})
                for pat, reason in KINSHI_PATTERNS:
                    for m in sorted(set(re.findall(pat, text))):
                        rows.append({"URL": url, "ページタイトル": meta["short_title"],
                                     "種別": "禁止表現", "発見テキスト": m,
                                     "推奨表記": f"【要確認】{reason}", "修正済み": ""})
                return rows

            findings: list = []
            done = 0
            with ThreadPoolExecutor(max_workers=workers) as ex:
                futs = {ex.submit(_content, u): u for u in urls}
                for f in as_completed(futs):
                    findings.extend(f.result())
                    done += 1
                    _upd(job_id,
                         phase=f"表記ゆれ確認中 ({done}/{n})",
                         progress=65 + round(done / n * 20))
                    time.sleep(REQUEST_DELAY)
            results["content"] = findings

        # ── Phase 4: Core Web Vitals ──────────────────────────────────────
        if "cwv" in check_types and n:
            import json as _j
            import urllib.request
            import urllib.parse as _up

            cwv: list = []
            for i, url in enumerate(urls):
                params = {"url": url, "strategy": strategy, "category": "performance"}
                if psi_key:
                    params["key"] = psi_key
                try:
                    req = urllib.request.Request(
                        f"{PSI_ENDPOINT}?{_up.urlencode(params)}",
                        headers={"User-Agent": "toyota-check/1.0"},
                    )
                    with urllib.request.urlopen(req, timeout=30) as r:
                        data = _j.loads(r.read().decode())
                    audits = data.get("lighthouseResult", {}).get("audits", {})
                    score = (data.get("lighthouseResult", {})
                             .get("categories", {}).get("performance", {}).get("score"))
                    def ms(k):
                        v = audits.get(k, {}).get("numericValue")
                        return round(v) if v else None
                    lcp = ms("largest-contentful-paint")
                    cls_ = round(audits.get("cumulative-layout-shift", {})
                                 .get("numericValue", 0), 3)
                    inp = ms("interaction-to-next-paint") or ms("total-blocking-time")
                    def rate(v, m):
                        t = THRESHOLDS.get(m, {})
                        if v is None:
                            return "N/A"
                        return "GOOD" if v <= t["good"] else ("NI" if v <= t["poor"] else "POOR")
                    ng = [m for m, v in [("LCP", lcp), ("CLS", cls_), ("INP", inp)]
                          if rate(v, m) in ("NI", "POOR")]
                    cwv.append({"URL": url,
                                "スコア": round(score * 100) if score else None,
                                "LCP(ms)": lcp, "LCP判定": rate(lcp, "LCP"),
                                "CLS": cls_, "CLS判定": rate(cls_, "CLS"),
                                "INP(ms)": inp, "INP判定": rate(inp, "INP"),
                                "KPI未達": ", ".join(ng) if ng else "✅"})
                except Exception as e:
                    cwv.append({"URL": url, "エラー": str(e)})
                _upd(job_id,
                     phase=f"CWV計測中 ({i + 1}/{n})",
                     progress=85 + round((i + 1) / n * 15))
                time.sleep(1.0)
            results["cwv"] = cwv

        _upd(job_id,
             status="done",
             phase="完了",
             progress=100,
             results=results,
             finished_at=datetime.now().isoformat())

    except Exception as e:
        import traceback
        _upd(job_id,
             status="error",
             phase=f"エラー: {e}",
             error=traceback.format_exc())
