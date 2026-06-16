"""toyota-check Streamlit Web アプリ"""

import io
import sys
import time
import pandas as pd
import streamlit as st
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# ─── ページ設定 ───────────────────────────────────────
st.set_page_config(
    page_title="toyota-check",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── パスワード認証 ───────────────────────────────────
def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    st.title("🚗 toyota-check")
    st.caption("toyota.jp 専用サイト検証ツール")
    st.divider()

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        pw = st.text_input("パスワード", type="password", key="pw_input",
                           placeholder="チーム共有パスワードを入力")
        if st.button("ログイン", use_container_width=True, type="primary"):
            correct = st.secrets.get("PASSWORD", "toyota-tqp-2026")
            if pw == correct:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("パスワードが違います")
    return False

if not check_password():
    st.stop()

# ─── インポート（認証後） ────────────────────────────
from fetcher import fetch_html, extract_meta
from config import CONCURRENT_WORKERS, REQUEST_DELAY

# ─── ヘルパー ─────────────────────────────────────────
def load_urls_from_excel(uploaded_file) -> list[str]:
    import openpyxl
    wb = openpyxl.load_workbook(uploaded_file, read_only=True)
    if "運用サイトマップ" not in wb.sheetnames:
        st.error("「運用サイトマップ」シートが見つかりません")
        return []
    ws = wb["運用サイトマップ"]
    urls = []
    for row in ws.iter_rows(min_row=6, values_only=True):
        u = row[7]
        if u and isinstance(u, str) and u.startswith("http"):
            urls.append(u.strip())
    return list(dict.fromkeys(urls))

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

# ─── UI ──────────────────────────────────────────────
st.title("🚗 toyota-check")
st.caption("toyota.jp 専用サイト検証ツール")

# サイドバー
with st.sidebar:
    st.header("設定")

    check_type = st.radio(
        "チェック種別",
        ["📄 タイトル取得", "🔗 リンクチェック", "⚡ Core Web Vitals", "📝 表記ゆれ・禁止表現"],
        index=0,
    )

    st.divider()
    st.subheader("URL ソース")
    url_source = st.radio("", ["Excelファイルをアップロード", "URLを直接入力"], label_visibility="collapsed")

    uploaded_file = None
    manual_urls = []

    if url_source == "Excelファイルをアップロード":
        uploaded_file = st.file_uploader(
            "tjpコンテンツ管理表.xlsx",
            type=["xlsx"],
            help="「運用サイトマップ」シートのフルURL列を読み込みます",
        )
        limit = st.slider("上限URL数（0=全件）", 0, 200, 0,
                          help="テスト時は10〜20程度に設定推奨")
    else:
        raw = st.text_area("URLを1行ずつ入力", height=200,
                           placeholder="https://toyota.jp/alphard/\nhttps://toyota.jp/prius/")
        manual_urls = [u.strip() for u in raw.splitlines() if u.strip().startswith("http")]
        limit = 0

    if check_type == "🔗 リンクチェック":
        toyota_only = st.checkbox(
            "toyota.jpリンクのみ",
            value=True,
            help="toyota.jpドメイン以外のリンクをチェック対象から除外（高速化）",
        )
    else:
        toyota_only = False

    if check_type == "⚡ Core Web Vitals":
        strategy = st.radio("計測デバイス", ["mobile", "desktop"], horizontal=True)
        psi_key = st.text_input("PSI API Key（任意）",
                                value=st.secrets.get("PSI_API_KEY", ""),
                                type="password",
                                help="未入力でも動作します（25,000回/日の上限あり）")
    else:
        strategy = "mobile"
        psi_key = st.secrets.get("PSI_API_KEY", "")

    st.divider()
    run_btn = st.button("▶ チェック実行", type="primary", use_container_width=True)

# ─── メインエリア ────────────────────────────────────
if not run_btn:
    st.info("サイドバーで設定を選び「チェック実行」を押してください。")

    with st.expander("📌 使い方"):
        st.markdown("""
| コマンド | 内容 |
|---|---|
| 📄 タイトル取得 | 全URLのtitle・descriptionを取得 |
| 🔗 リンクチェック | 404・リンク切れを検出 |
| ⚡ Core Web Vitals | LCP/CLS/INP をTQP KPIと照合 |
| 📝 表記ゆれ・禁止表現 | 景表法・誇大表現・表記ゆれを検出 |

**TQP KPI 閾値**
- LCP ≤ 2.5s　CLS ≤ 0.1　INP ≤ 200ms
        """)
    st.stop()

# ─── URL 収集 ────────────────────────────────────────
if url_source == "Excelファイルをアップロード":
    if not uploaded_file:
        st.warning("Excelファイルをアップロードしてください。")
        st.stop()
    with st.spinner("Excelを読み込み中..."):
        urls = load_urls_from_excel(uploaded_file)
    if limit:
        urls = urls[:limit]
else:
    urls = manual_urls

if not urls:
    st.error("URLが見つかりません。")
    st.stop()

st.success(f"対象URL: **{len(urls)} 件**")

# ─── 各チェック実行 ───────────────────────────────────

# ════════════════════════════════════════
# 📄 タイトル取得
# ════════════════════════════════════════
if check_type == "📄 タイトル取得":
    from concurrent.futures import ThreadPoolExecutor, as_completed

    results = []
    progress = st.progress(0, text="取得中...")
    status_box = st.empty()

    def _fetch(url):
        code, html = fetch_html(url)
        if code == 200 and html:
            meta = extract_meta(html)
            return {"URL": url, "ステータス": code,
                    "タイトル": meta["short_title"],
                    "description": meta["description"][:100],
                    "判定": "✅" if meta["short_title"] else "⚠️ titleなし"}
        return {"URL": url, "ステータス": code, "タイトル": "", "description": "", "判定": "❌"}

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futures = {ex.submit(_fetch, u): u for u in urls}
        done = 0
        for future in as_completed(futures):
            results.append(future.result())
            done += 1
            progress.progress(done / len(urls), text=f"取得中... {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)

    progress.empty()
    df = pd.DataFrame(results)
    ng = df[df["判定"] != "✅"]

    col1, col2 = st.columns(2)
    col1.metric("✅ 取得成功", len(df[df["判定"] == "✅"]))
    col2.metric("⚠️ 要確認", len(ng))

    st.subheader("結果一覧")
    st.dataframe(df, use_container_width=True, height=500)

    st.download_button("📥 Excelダウンロード", to_excel_bytes(df),
                       "titles_result.xlsx", use_container_width=True)

# ════════════════════════════════════════
# 🔗 リンクチェック
# ════════════════════════════════════════
elif check_type == "🔗 リンクチェック":
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from fetcher import extract_resources
    from urllib.parse import urlparse

    # Phase 1: 各ページのリソースを収集
    st.caption("Phase 1/2 — ページ内リソースを収集中（リンク・CSS・JS・画像・メディア）")
    progress1 = st.progress(0, text="収集中...")
    page_resources: dict = {}  # {src_url: {"status": int, "resources": [(url, type), ...]}}

    def _collect(url):
        code, html = fetch_html(url)
        resources = extract_resources(html, url) if (code == 200 and html) else []
        if toyota_only:
            resources = [(u, t) for u, t in resources if urlparse(u).netloc == "toyota.jp"]
        return url, code, resources

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futs = {ex.submit(_collect, u): u for u in urls}
        done = 0
        for f in as_completed(futs):
            src, code, resources = f.result()
            page_resources[src] = {"status": code, "resources": resources}
            done += 1
            progress1.progress(done / len(urls), text=f"リソース収集中... {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)
    progress1.empty()

    all_resource_urls = list({u for d in page_resources.values() for u, _ in d["resources"]})
    st.caption(f"Phase 2/2 — 発見リソース {len(all_resource_urls)} 件のステータスを確認中")

    # Phase 2: 収集したリソースのステータスチェック
    progress2 = st.progress(0, text="チェック中...")
    resource_status: dict = {}

    def _ping(url):
        code, _ = fetch_html(url)
        return url, code

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futs = {ex.submit(_ping, u): u for u in all_resource_urls}
        done = 0
        for f in as_completed(futs):
            url, code = f.result()
            resource_status[url] = code
            done += 1
            progress2.progress(done / len(all_resource_urls), text=f"確認中... {done}/{len(all_resource_urls)}")
            time.sleep(REQUEST_DELAY)
    progress2.empty()

    # Phase 3: 発見ページ × リソースURL で集計
    broken_rows = []
    all_rows = []
    for src in sorted(page_resources):
        d = page_resources[src]
        for res_url, res_type in d["resources"]:
            code = resource_status.get(res_url, 0)
            broken = code in (404, 410) or code == 0
            row = {
                "発見ページ": src,
                "リソースURL": res_url,
                "種別": res_type,
                "ドメイン": urlparse(res_url).netloc,
                "ステータス": code,
                "判定": "❌ 切れ" if broken else "✅ 正常",
            }
            all_rows.append(row)
            if broken:
                broken_rows.append(row)

    broken_pages = len({r["発見ページ"] for r in broken_rows})
    col1, col2, col3 = st.columns(3)
    col1.metric("対象ページ数", len(urls))
    col2.metric("❌ 問題ページ数", broken_pages)
    col3.metric("❌ 問題リソース件数", len(broken_rows))

    if broken_rows:
        st.error(f"問題リソース: {len(broken_rows)} 件（{broken_pages} ページで発見）")
        st.subheader("❌ 問題リソース一覧")
        df_broken = pd.DataFrame(broken_rows)
        # 種別フィルタ
        types = ["すべて"] + sorted(df_broken["種別"].unique().tolist())
        sel = st.selectbox("種別フィルタ", types)
        disp = df_broken if sel == "すべて" else df_broken[df_broken["種別"] == sel]
        st.dataframe(disp, use_container_width=True, height=400)
        st.download_button("📥 問題リソースExcelダウンロード", to_excel_bytes(df_broken),
                           "links_broken.xlsx", use_container_width=True)
    else:
        st.success("問題のあるリソースは検出されませんでした ✅")

    with st.expander(f"全リソース一覧（{len(all_rows)} 件）"):
        df_all = pd.DataFrame(all_rows)
        st.dataframe(df_all, use_container_width=True, height=400)
        st.download_button("📥 全件Excelダウンロード", to_excel_bytes(df_all),
                           "links_all.xlsx", use_container_width=True, key="dl_all")

# ════════════════════════════════════════
# ⚡ Core Web Vitals
# ════════════════════════════════════════
elif check_type == "⚡ Core Web Vitals":
    import json, urllib.request, urllib.parse
    from config import PSI_ENDPOINT, THRESHOLDS

    results = []
    progress = st.progress(0, text="計測中...")

    def _cwv(url):
        params = {"url": url, "strategy": strategy, "category": "performance"}
        if psi_key:
            params["key"] = psi_key
        api_url = f"{PSI_ENDPOINT}?{urllib.parse.urlencode(params)}"
        try:
            req = urllib.request.Request(api_url, headers={"User-Agent": "toyota-check/1.0"})
            with urllib.request.urlopen(req, timeout=30) as res:
                data = json.loads(res.read().decode())
        except Exception as e:
            return {"URL": url, "エラー": str(e)}

        audits = data.get("lighthouseResult", {}).get("audits", {})
        score = data.get("lighthouseResult", {}).get("categories", {}).get("performance", {}).get("score")

        def ms(k): v = audits.get(k, {}).get("numericValue"); return round(v) if v else None
        lcp = ms("largest-contentful-paint")
        cls = round(audits.get("cumulative-layout-shift", {}).get("numericValue", 0), 3)
        inp = ms("interaction-to-next-paint") or ms("total-blocking-time")

        def rate(v, m):
            t = THRESHOLDS.get(m, {})
            if v is None: return "N/A"
            return "GOOD" if v <= t["good"] else ("NI" if v <= t["poor"] else "POOR")

        ng = [m for m, v in [("LCP", lcp), ("CLS", cls), ("INP", inp)]
              if rate(v, m) in ("NI", "POOR")]
        return {
            "URL": url, "スコア": round(score * 100) if score else None,
            "LCP(ms)": lcp, "LCP判定": rate(lcp, "LCP"),
            "CLS": cls,     "CLS判定": rate(cls, "CLS"),
            "INP(ms)": inp, "INP判定": rate(inp, "INP"),
            "KPI未達": ", ".join(ng) if ng else "✅",
        }

    for i, url in enumerate(urls):
        results.append(_cwv(url))
        progress.progress((i + 1) / len(urls), text=f"計測中... {i+1}/{len(urls)}")
        time.sleep(1.0)

    progress.empty()
    df = pd.DataFrame(results)
    ng_df = df[df["KPI未達"] != "✅"] if "KPI未達" in df.columns else pd.DataFrame()

    col1, col2 = st.columns(2)
    col1.metric("✅ KPI達成", len(df) - len(ng_df))
    col2.metric("❌ KPI未達", len(ng_df))

    if len(ng_df):
        st.subheader("❌ KPI未達ページ")
        st.dataframe(ng_df, use_container_width=True)

    with st.expander("全件表示"):
        st.dataframe(df, use_container_width=True, height=400)

    st.download_button("📥 Excelダウンロード", to_excel_bytes(df),
                       "cwv_result.xlsx", use_container_width=True)

# ════════════════════════════════════════
# 📝 表記ゆれ・禁止表現
# ════════════════════════════════════════
elif check_type == "📝 表記ゆれ・禁止表現":
    import re
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from fetcher import extract_text_blocks
    from config import HYOKI_YURE, KINSHI_PATTERNS

    results = []
    progress = st.progress(0, text="解析中...")

    def _content(url):
        code, html = fetch_html(url)
        if code != 200 or not html:
            return {"URL": url, "ステータス": code, "表記ゆれ": "", "禁止表現": "", "判定": "❌"}
        text = " ".join(extract_text_blocks(html))
        yure = [f"{list(set(re.findall(p, text, re.I)))[:2]}→{c}"
                for p, c in HYOKI_YURE.items() if re.search(p, text, re.I)]
        kinshi = [f"{list(set(re.findall(p, text)))[:2]}({r})"
                  for p, r in KINSHI_PATTERNS if re.search(p, text)]
        return {
            "URL": url, "ステータス": code,
            "表記ゆれ": " / ".join(yure),
            "禁止表現": " / ".join(kinshi),
            "判定": "⚠️ 要確認" if (yure or kinshi) else "✅",
        }

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futures = {ex.submit(_content, u): u for u in urls}
        done = 0
        for future in as_completed(futures):
            results.append(future.result())
            done += 1
            progress.progress(done / len(urls), text=f"解析中... {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)

    progress.empty()
    df = pd.DataFrame(results)
    ng_df = df[df["判定"] == "⚠️ 要確認"]

    col1, col2 = st.columns(2)
    col1.metric("✅ 問題なし", len(df[df["判定"] == "✅"]))
    col2.metric("⚠️ 要確認", len(ng_df))

    if len(ng_df):
        st.subheader("⚠️ 要確認ページ")
        st.dataframe(ng_df, use_container_width=True)

    with st.expander("全件表示"):
        st.dataframe(df, use_container_width=True, height=400)

    st.download_button("📥 Excelダウンロード", to_excel_bytes(df),
                       "content_result.xlsx", use_container_width=True)
