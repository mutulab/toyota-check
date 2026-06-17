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
        ["📄 タイトル取得", "🔗 リンクチェック", "⚡ Core Web Vitals",
         "📝 表記ゆれ・禁止表現", "🔍 アプリ・機能検出", "🕷️ バックグラウンドクロール"],
        index=0,
    )

    # ── バックグラウンドクロール専用設定 ──────────────────────────────
    bg_cfg: dict = {}
    if check_type == "🕷️ バックグラウンドクロール":
        st.divider()
        st.subheader("ジョブ設定")

        _bg_url_src = st.radio(
            "URL取得方法",
            ["🕷️ クロール（自動収集）", "📊 コンテンツ管理票（Excel）"],
            key="bg_url_src",
        )
        if _bg_url_src.startswith("🕷"):
            bg_cfg["url_source_type"] = "crawl"
            bg_cfg["start_url"] = st.text_input("開始URL", value="https://toyota.jp/")
            bg_cfg["max_pages"] = st.slider("最大収集ページ数", 10, 1000, 100)
            bg_cfg["depth"]     = int(st.number_input("クロール深さ", 1, 4, 2))
        else:
            bg_cfg["url_source_type"] = "excel"
            bg_cfg["start_url"] = ""
            _bg_exc = st.file_uploader(
                "コンテンツ管理票.xlsx", type=["xlsx"], key="bg_exc_ul"
            )
            if _bg_exc:
                _fk = f"{_bg_exc.name}_{_bg_exc.size}"
                if st.session_state.get("_bg_exc_key") != _fk:
                    _exc_urls = load_urls_from_excel(_bg_exc)
                    st.session_state["bg_excel_urls"] = _exc_urls
                    st.session_state["_bg_exc_key"] = _fk
            bg_cfg["urls"]      = st.session_state.get("bg_excel_urls", [])
            bg_cfg["max_pages"] = len(bg_cfg["urls"])
            bg_cfg["depth"]     = 0
            if bg_cfg["urls"]:
                st.caption(f"✅ {len(bg_cfg['urls'])} 件のURLを読み込みました")
            else:
                st.info("コンテンツ管理票をアップロードしてください")

        bg_cfg["toyota_only"]  = st.checkbox("toyota.jpリンクのみ", value=True)
        _bg_checks = st.multiselect(
            "実行するチェック",
            ["リンクチェック", "表記ゆれ・禁止表現", "Core Web Vitals", "アプリ・機能検出"],
            default=["リンクチェック"],
        )
        _map = {"リンクチェック": "link", "表記ゆれ・禁止表現": "content",
                "Core Web Vitals": "cwv", "アプリ・機能検出": "app"}
        bg_cfg["check_types"] = [_map[c] for c in _bg_checks]
        if "link" in bg_cfg["check_types"]:
            st.caption("リンクチェック: リソース種別")
            c1, c2 = st.columns(2)
            _sel: set = set()
            if c1.checkbox("リンク(<a>)",   value=True,  key="bg_rl"): _sel.add("リンク")
            if c1.checkbox("CSS",           value=True,  key="bg_rc"): _sel.add("CSS/スタイル")
            if c1.checkbox("JavaScript",    value=True,  key="bg_rj"): _sel.add("JavaScript")
            if c2.checkbox("画像",          value=True,  key="bg_ri"): _sel.update({"画像", "画像(インラインCSS)"})
            if c2.checkbox("メディア",      value=False, key="bg_rm"): _sel.add("メディア")
            if c2.checkbox("iframe",        value=False, key="bg_rf"): _sel.add("iframe")
            if st.checkbox("CSS内リソース", value=False, key="bg_rcr"): _sel.update({"画像(CSS)", "フォント(CSS)", "CSSリソース"})
            bg_cfg["selected_res"] = list(_sel)
        if "content" in bg_cfg["check_types"]:
            bg_cfg["custom_dict"] = st.text_area(
                "カスタム辞書（誤表記|推奨表記）", height=80,
                placeholder="ウエブサイト|WEBサイト", key="bg_dict",
            )
        if "cwv" in bg_cfg["check_types"]:
            bg_cfg["strategy"] = st.radio("CWVデバイス", ["mobile", "desktop"],
                                          horizontal=True, key="bg_strat")
            bg_cfg["psi_key"]  = st.text_input("PSI API Key", type="password",
                                               value=st.secrets.get("PSI_API_KEY", ""),
                                               key="bg_psi")
        # URL source 変数をデフォルト値で初期化（後続コードで参照されるため）
        url_source = None
        uploaded_file = None
        manual_urls: list = []
        crawl_start = ""
        crawl_max = 100
        crawl_depth = 2
        limit = 0
        toyota_only = False
        selected_res_types: set = set()
        custom_dict_raw = ""
        strategy = "mobile"
        psi_key = st.secrets.get("PSI_API_KEY", "")

    else:
        # ── 通常モード: URL ソース ──────────────────────────────────────
        st.divider()
        st.subheader("URL ソース")
        url_source = st.radio(
            "",
            ["Excelファイルをアップロード", "URLを直接入力", "🕷️ クロール（自動収集）"],
            label_visibility="collapsed",
        )

        uploaded_file = None
        manual_urls = []
        crawl_start = ""
        crawl_max = 100
        crawl_depth = 2

        if url_source == "Excelファイルをアップロード":
            uploaded_file = st.file_uploader(
                "tjpコンテンツ管理表.xlsx",
                type=["xlsx"],
                help="「運用サイトマップ」シートのフルURL列を読み込みます",
            )
            limit = st.slider("上限URL数（0=全件）", 0, 200, 0,
                              help="テスト時は10〜20程度に設定推奨")
        elif url_source == "URLを直接入力":
            raw = st.text_area("URLを1行ずつ入力", height=200,
                               placeholder="https://toyota.jp/alphard/\nhttps://toyota.jp/prius/")
            manual_urls = [u.strip() for u in raw.splitlines() if u.strip().startswith("http")]
            limit = 0
        else:  # クロール（自動収集）
            crawl_start = st.text_input("開始URL", value="https://toyota.jp/")
            crawl_max = st.slider("最大収集ページ数", 10, 1000, 100,
                                  help="まず50前後で試してください")
            crawl_depth = st.number_input("クロール深さ（階層数）", min_value=1, max_value=4, value=2,
                                          help="1=直リンクのみ / 2=その先も辿る")
            limit = 0

        # ── リンクチェック: リソース種別選択 ───────────────────────────
        if check_type == "🔗 リンクチェック":
            toyota_only = st.checkbox(
                "toyota.jpリンクのみ",
                value=True,
                help="toyota.jpドメイン以外を除外（高速化）",
            )
            st.divider()
            st.caption("チェック対象リソース種別")
            c1, c2 = st.columns(2)
            _sel2: set = set()
            if c1.checkbox("リンク(<a>)",  value=True,  key="rl"): _sel2.add("リンク")
            if c1.checkbox("CSS",          value=True,  key="rc"): _sel2.add("CSS/スタイル")
            if c1.checkbox("JavaScript",   value=True,  key="rj"): _sel2.add("JavaScript")
            if c2.checkbox("画像",         value=True,  key="ri"): _sel2.update({"画像", "画像(インラインCSS)"})
            if c2.checkbox("メディア",     value=False, key="rm"): _sel2.add("メディア")
            if c2.checkbox("iframe",       value=False, key="rf"): _sel2.add("iframe")
            if st.checkbox("CSS内リソース（背景画像・フォント）", value=False, key="rcr"):
                _sel2.update({"画像(CSS)", "フォント(CSS)", "CSSリソース"})
            selected_res_types = _sel2
        else:
            toyota_only = False
            selected_res_types = set()

        if check_type == "📝 表記ゆれ・禁止表現":
            st.divider()
            st.subheader("表記ゆれ辞書")
            from dict_loader import load_flat, source_label as _src_label
            _cur_flat = load_flat()
            st.caption(f"{_src_label()} · {len(_cur_flat)} 件")
            with st.expander("辞書を更新（Excelアップロード）"):
                _dict_xl = st.file_uploader(
                    "toyota_lexus用語リスト.xlsx", type=["xlsx"], key="dict_xl_up"
                )
                if _dict_xl:
                    _dk = f"{_dict_xl.name}_{_dict_xl.size}"
                    if st.session_state.get("_dict_xl_key") != _dk:
                        st.session_state["_dict_new_bytes"] = _dict_xl.read()
                        st.session_state["_dict_xl_key"] = _dk
            custom_dict_raw = st.text_area(
                "追加エントリ（1行1件: NG表記|推奨表記）",
                height=100,
                placeholder="ウエブサイト|WEBサイト\nお問合せ|お問い合わせ",
                help="辞書にない追加チェック項目。正規表現も使用可。",
            )
        else:
            custom_dict_raw = ""

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

# ─── バックグラウンドクロール UI（常時表示 / 通常フローをスキップ） ──
if check_type == "🕷️ バックグラウンドクロール":
    try:
        import crawler as _cw
    except Exception as _e:
        st.error(f"crawler モジュールの読み込みに失敗しました: {_e}")
        st.stop()

    tab_new, tab_status = st.tabs(["▶ 新規ジョブ開始", "🔍 ジョブ確認"])

    with tab_new:
        st.subheader("ジョブ設定")
        st.json({k: v for k, v in bg_cfg.items() if k != "psi_key"})
        if run_btn:  # サイドバーの「▶ チェック実行」= ジョブ開始
            if bg_cfg.get("url_source_type") == "excel" and not bg_cfg.get("urls"):
                st.error("コンテンツ管理票のExcelをアップロードしてください。")
            else:
                try:
                    jid = _cw.start_job(bg_cfg)
                    st.session_state["bg_job_id"] = jid
                    st.success(f"ジョブを開始しました。**ジョブID: `{jid}`**")
                    st.info("ブラウザを閉じても処理は継続します。「🔍 ジョブ確認」タブで結果を確認してください。")
                except Exception as _e:
                    st.error(f"ジョブ開始に失敗しました: {_e}")
        else:
            st.info("設定を確認して、サイドバーの「▶ チェック実行」を押すとジョブが開始されます。")

    with tab_status:
        # ── ジョブ選択（履歴 or 直接入力） ───────────────────────────
        recent = _cw.list_jobs()
        job_id_input = ""

        if recent:
            _NONE = "— IDを直接入力 —"
            def _jlabel(j):
                src = j["cfg"].get("start_url") or f"Excel ({j['cfg'].get('max_pages', '?')}件)"
                return f"{j['id']}  [{j['status']}]  {src}  ({j['started_at'][:16]})"
            _opts: dict = {_jlabel(j): j["id"] for j in recent}
            _all_keys = [_NONE] + list(_opts.keys())
            _sess_id = st.session_state.get("bg_job_id", "")
            _def_key = next((k for k, v in _opts.items() if v == _sess_id), _NONE)
            _sel = st.selectbox(
                "📋 ジョブ履歴から選択",
                _all_keys,
                index=_all_keys.index(_def_key),
            )
            if _sel != _NONE:
                job_id_input = _opts[_sel]

        if not job_id_input:
            job_id_input = st.text_input(
                "ジョブIDを直接入力",
                value=st.session_state.get("bg_job_id", ""),
                placeholder="例: A1B2C3D4",
            ).strip().upper()

        # ── ジョブ詳細 ────────────────────────────────────────────────
        if job_id_input:
            job = _cw.get_job(job_id_input)
            if not job:
                st.error("ジョブが見つかりません。IDを確認してください。")
            else:
                st.progress(job["progress"] / 100, text=job["phase"])
                col1, col2, col3 = st.columns(3)
                col1.metric("ステータス", job["status"])
                col2.metric("収集ページ数", job.get("url_count", 0))
                col3.metric("進捗", f"{job['progress']}%")
                st.caption(f"開始: {job['started_at']}　完了: {job.get('finished_at') or '—'}")

                if job["status"] == "error":
                    st.error(f"エラー: {job.get('error', '')}")

                if job["status"] == "running":
                    from datetime import datetime as _dt
                    _last = job.get("last_updated_at") or job.get("started_at", "")
                    try:
                        _elapsed = (_dt.now() - _dt.fromisoformat(_last)).total_seconds()
                        if _elapsed > 600:
                            st.warning(
                                f"⚠️ 最終更新から {int(_elapsed / 60)} 分経過しています。"
                                "タスクがスタックしている可能性があります。"
                            )
                    except Exception:
                        pass
                    st.info("実行中です。完了したら「🔄 更新」を押してください。")
                    if st.button("🔄 更新"):
                        st.rerun()

                if job["status"] == "done":
                    res = job.get("results", {})
                    if "link" in res:
                        df_l = pd.DataFrame(res["link"]) if res["link"] else pd.DataFrame()
                        st.subheader(f"🔗 リンク切れ ({len(df_l)} 件)")
                        if not df_l.empty:
                            st.dataframe(df_l, use_container_width=True, height=300)
                            c1, c2 = st.columns(2)
                            c1.download_button("📥 Excel", to_excel_bytes(df_l),
                                               f"links_{job_id_input}.xlsx", key="bg_dl_lx")
                            c2.download_button("📥 CSV",
                                               df_l.to_csv(index=False).encode("utf-8-sig"),
                                               f"links_{job_id_input}.csv", mime="text/csv",
                                               key="bg_dl_lc")
                    if "content" in res:
                        df_c = pd.DataFrame(res["content"]) if res["content"] else pd.DataFrame()
                        st.subheader(f"📝 表記ゆれ・禁止表現 ({len(df_c)} 件)")
                        if not df_c.empty:
                            st.dataframe(df_c, use_container_width=True, height=300)
                            c1, c2 = st.columns(2)
                            c1.download_button("📥 指示書 Excel", to_excel_bytes(df_c),
                                               f"content_{job_id_input}.xlsx", key="bg_dl_cx")
                            c2.download_button("📥 指示書 CSV",
                                               df_c.to_csv(index=False).encode("utf-8-sig"),
                                               f"content_{job_id_input}.csv", mime="text/csv",
                                               key="bg_dl_cc")
                    if "cwv" in res:
                        df_w = pd.DataFrame(res["cwv"]) if res["cwv"] else pd.DataFrame()
                        st.subheader(f"⚡ Core Web Vitals ({len(df_w)} 件)")
                        if not df_w.empty:
                            st.dataframe(df_w, use_container_width=True, height=300)
                            st.download_button("📥 CWV Excel", to_excel_bytes(df_w),
                                               f"cwv_{job_id_input}.xlsx", key="bg_dl_wx")
                    if "app" in res:
                        df_a = pd.DataFrame(res["app"]) if res["app"] else pd.DataFrame()
                        app_cnt = len(df_a[df_a["アプリ性スコア"] > 0]) if not df_a.empty else 0
                        st.subheader(f"🔍 アプリ・機能検出 (機能あり: {app_cnt} / {len(df_a)} ページ)")
                        if not df_a.empty:
                            st.dataframe(df_a, use_container_width=True, height=300)
                            st.download_button("📥 アプリ検出 Excel", to_excel_bytes(df_a),
                                               f"app_{job_id_input}.xlsx", key="bg_dl_ax")
    st.stop()

# ─── 表記ゆれ辞書差分確認（常時表示） ────────────────────────────────
if check_type == "📝 表記ゆれ・禁止表現" and "_dict_new_bytes" in st.session_state:
    import io as _io
    from dict_loader import load_flat, read_excel_flat, save_override
    try:
        _new_flat = read_excel_flat(_io.BytesIO(st.session_state["_dict_new_bytes"]))
        _cur_flat = load_flat()
        _added   = {k: v for k, v in _new_flat.items() if k not in _cur_flat}
        _removed = {k: v for k, v in _cur_flat.items() if k not in _new_flat}
        _changed = {k: (_cur_flat[k], v) for k, v in _new_flat.items()
                    if k in _cur_flat and _cur_flat[k] != v}
        total_delta = len(_added) + len(_removed) + len(_changed)
        if total_delta:
            st.subheader(f"📋 辞書差分プレビュー（+{len(_added)} / -{len(_removed)} / ~{len(_changed)}）")
            if _added:
                with st.expander(f"✅ 追加: {len(_added)} 件"):
                    st.dataframe(
                        pd.DataFrame([{"NG表記": k, "推奨表記": v} for k, v in _added.items()]),
                        use_container_width=True,
                    )
            if _removed:
                with st.expander(f"❌ 削除: {len(_removed)} 件"):
                    st.dataframe(
                        pd.DataFrame([{"NG表記": k, "推奨表記": v} for k, v in _removed.items()]),
                        use_container_width=True,
                    )
            if _changed:
                with st.expander(f"✏️ 変更: {len(_changed)} 件"):
                    st.dataframe(
                        pd.DataFrame([{"NG表記": k, "旧推奨": old, "新推奨": new}
                                      for k, (old, new) in _changed.items()]),
                        use_container_width=True,
                    )
            _ca, _cb = st.columns(2)
            if _ca.button("✅ 差分を適用", type="primary", key="apply_dict"):
                save_override(st.session_state["_dict_new_bytes"])
                for _k in ("_dict_new_bytes", "_dict_xl_key"):
                    st.session_state.pop(_k, None)
                st.success("辞書を更新しました。")
                st.rerun()
            if _cb.button("❌ キャンセル", key="cancel_dict"):
                for _k in ("_dict_new_bytes", "_dict_xl_key"):
                    st.session_state.pop(_k, None)
                st.rerun()
        else:
            st.info("アップロードした辞書と現在の辞書に差分はありません。")
            for _k in ("_dict_new_bytes", "_dict_xl_key"):
                st.session_state.pop(_k, None)
    except Exception as _dict_err:
        st.error(f"辞書の読み込みに失敗しました: {_dict_err}")
        for _k in ("_dict_new_bytes", "_dict_xl_key"):
            st.session_state.pop(_k, None)

# ─── 通常モード: ボタンを押すまで待機 ──────────────────────────────
if not run_btn:
    st.info("サイドバーで設定を入力し「▶ チェック実行」を押してください。")
    st.stop()

# ─── URL 収集（通常モード） ──────────────────────────────────────────
if url_source == "Excelファイルをアップロード":
    if not uploaded_file:
        st.warning("Excelファイルをアップロードしてください。")
        st.stop()
    with st.spinner("Excelを読み込み中..."):
        urls = load_urls_from_excel(uploaded_file)
    if limit:
        urls = urls[:limit]

elif url_source == "URLを直接入力":
    urls = manual_urls

else:  # 🕷️ クロール
    from urllib.parse import urlparse
    from collections import deque
    from fetcher import extract_links as _extract_links

    if not crawl_start or not crawl_start.startswith("http"):
        st.error("クロール開始URLを入力してください。")
        st.stop()

    base_domain = urlparse(crawl_start).netloc
    visited: set = set()
    queue: deque = deque([(crawl_start.rstrip("/"), 0)])
    urls: list = []

    st.info(f"**{crawl_start}** からクロール中（最大 {crawl_max} ページ、深さ {crawl_depth}）")
    crawl_progress = st.progress(0, text="クロール中...")
    crawl_status = st.empty()

    while queue and len(urls) < crawl_max:
        url, depth = queue.popleft()
        norm = url.rstrip("/").split("?")[0].split("#")[0]
        if norm in visited:
            continue
        visited.add(norm)

        code, html = fetch_html(url)
        if code == 200 and html:
            urls.append(url)
            crawl_progress.progress(
                min(len(urls) / crawl_max, 1.0),
                text=f"収集中... {len(urls)} / {crawl_max} ページ",
            )
            crawl_status.caption(f"→ {url}")

            if depth < crawl_depth:
                for link in _extract_links(html, url):
                    if urlparse(link).netloc == base_domain:
                        clean = link.rstrip("/").split("?")[0].split("#")[0]
                        if clean not in visited:
                            queue.append((link, depth + 1))

        time.sleep(REQUEST_DELAY)

    crawl_progress.empty()
    crawl_status.empty()
    st.success(f"クロール完了: **{len(urls)} ページ**を収集しました")

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
    from fetcher import extract_resources, extract_css_urls
    from urllib.parse import urlparse

    # Phase 1: 各ページのリソースを収集
    st.caption("Phase 1/2 — ページ内リソースを収集中（リンク・CSS・JS・画像・メディア・CSSリソース）")
    progress1 = st.progress(0, text="収集中...")
    page_resources: dict = {}  # {src_url: {"status": int, "resources": [(url, type), ...]}}

    _css_res_types = {"画像(CSS)", "フォント(CSS)", "CSSリソース"}

    def _collect(url):
        code, html = fetch_html(url)
        resources = extract_resources(html, url) if (code == 200 and html) else []

        # CSS内リソースが選択対象の場合のみCSSファイルをフェッチ
        if not selected_res_types or selected_res_types & _css_res_types:
            for css_url in [u for u, t in resources if t == "CSS/スタイル"]:
                css_code, css_text = fetch_html(css_url)
                if css_code == 200 and css_text:
                    resources.extend(extract_css_urls(css_text, css_url))

        if selected_res_types:
            resources = [(u, t) for u, t in resources if t in selected_res_types]
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
        c1, c2 = st.columns(2)
        c1.download_button("📥 Excel", to_excel_bytes(df_broken),
                           "links_broken.xlsx", use_container_width=True, key="dl_broken_xlsx")
        c2.download_button("📥 CSV", df_broken.to_csv(index=False).encode("utf-8-sig"),
                           "links_broken.csv", mime="text/csv",
                           use_container_width=True, key="dl_broken_csv")
    else:
        st.success("問題のあるリソースは検出されませんでした ✅")

    with st.expander(f"全リソース一覧（{len(all_rows)} 件）"):
        df_all = pd.DataFrame(all_rows)
        st.dataframe(df_all, use_container_width=True, height=400)
        ca, cb = st.columns(2)
        ca.download_button("📥 全件Excel", to_excel_bytes(df_all),
                           "links_all.xlsx", use_container_width=True, key="dl_all_xlsx")
        cb.download_button("📥 全件CSV", df_all.to_csv(index=False).encode("utf-8-sig"),
                           "links_all.csv", mime="text/csv",
                           use_container_width=True, key="dl_all_csv")

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
    from fetcher import extract_meta, extract_text_blocks
    from config import KINSHI_PATTERNS
    from dict_loader import load_for_check

    # Excel辞書 + カスタム追加エントリをマージ
    merged_dict = load_for_check()
    for line in custom_dict_raw.splitlines():
        line = line.strip()
        if "|" in line:
            wrong, correct = line.split("|", 1)
            merged_dict[re.escape(wrong.strip())] = correct.strip()

    results = []
    progress = st.progress(0, text="解析中...")

    def _content(url):
        code, html = fetch_html(url)
        if code != 200 or not html:
            return {"url": url, "code": code, "title": "", "findings": [], "判定": "❌ 取得失敗"}
        meta = extract_meta(html)
        text = " ".join(extract_text_blocks(html))
        findings = []

        for pattern, recommended in merged_dict.items():
            for match in sorted(set(re.findall(pattern, text, re.I))):
                findings.append({
                    "URL": url,
                    "ページタイトル": meta["short_title"],
                    "種別": "表記ゆれ",
                    "発見テキスト": match,
                    "推奨表記": recommended,
                    "修正済み": "",
                })

        for pattern, reason in KINSHI_PATTERNS:
            for match in sorted(set(re.findall(pattern, text))):
                findings.append({
                    "URL": url,
                    "ページタイトル": meta["short_title"],
                    "種別": "禁止表現",
                    "発見テキスト": match,
                    "推奨表記": f"【要確認】{reason}",
                    "修正済み": "",
                })

        return {
            "url": url, "code": code, "title": meta["short_title"],
            "findings": findings,
            "判定": "⚠️ 要確認" if findings else "✅",
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

    # ページ単位サマリ
    summary_rows = [{
        "URL": r["url"],
        "ページタイトル": r["title"],
        "ステータス": r["code"],
        "表記ゆれ件数": sum(1 for f in r["findings"] if f["種別"] == "表記ゆれ"),
        "禁止表現件数": sum(1 for f in r["findings"] if f["種別"] == "禁止表現"),
        "判定": r["判定"],
    } for r in results]

    # 指示書（1指摘1行）
    all_findings = sorted(
        [f for r in results for f in r["findings"]],
        key=lambda x: (x["種別"], x["URL"])
    )
    instruction_rows = [{"No.": i, **f} for i, f in enumerate(all_findings, 1)]

    df_summary = pd.DataFrame(summary_rows)
    ng_count = len([r for r in results if r["判定"] == "⚠️ 要確認"])

    col1, col2, col3 = st.columns(3)
    col1.metric("✅ 問題なし", len(results) - ng_count)
    col2.metric("⚠️ 要確認ページ", ng_count)
    col3.metric("📋 指摘件数合計", len(instruction_rows))

    st.subheader("ページ単位サマリ")
    st.dataframe(df_summary, use_container_width=True, height=350)

    if instruction_rows:
        st.subheader("📋 修正指示書")
        df_inst = pd.DataFrame(instruction_rows)

        kinds = ["すべて"] + sorted(df_inst["種別"].unique().tolist())
        sel_kind = st.selectbox("種別フィルタ", kinds)
        disp_inst = df_inst if sel_kind == "すべて" else df_inst[df_inst["種別"] == sel_kind]
        st.dataframe(disp_inst, use_container_width=True, height=420)

        c1, c2 = st.columns(2)
        c1.download_button(
            "📥 指示書 Excel",
            to_excel_bytes(df_inst),
            "修正指示書.xlsx",
            use_container_width=True,
        )
        c2.download_button(
            "📥 指示書 CSV",
            df_inst.to_csv(index=False).encode("utf-8-sig"),
            "修正指示書.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_inst_csv",
        )
    else:
        st.success("問題のある表現は検出されませんでした ✅")

# ════════════════════════════════════════
# 🔍 アプリ・機能検出
# ════════════════════════════════════════
elif check_type == "🔍 アプリ・機能検出":
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from fetcher import extract_meta
    from detector import summarize, CATEGORIES

    rows = []
    progress = st.progress(0, text="解析中...")

    def _detect(url):
        code, html = fetch_html(url)
        if code != 200 or not html:
            return {"URL": url, "タイトル": "", "アプリ性スコア": 0, "検出機能": f"❌ HTTP {code}",
                    **{c: "" for c in CATEGORIES}}
        meta = extract_meta(html)
        return summarize(url, html, meta["short_title"])

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futs = {ex.submit(_detect, u): u for u in urls}
        done = 0
        for f in as_completed(futs):
            rows.append(f.result())
            done += 1
            progress.progress(done / len(urls), text=f"解析中... {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)

    progress.empty()
    df_app = pd.DataFrame(rows)

    # ── サマリーメトリクス ────────────────────────────────────────────
    n_with_feat = len(df_app[df_app["アプリ性スコア"] > 0])
    n_high      = len(df_app[df_app["アプリ性スコア"] >= 3])
    cat_counts  = {c: int((df_app[c] != "").sum()) for c in CATEGORIES}

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("解析ページ数", len(df_app))
    c2.metric("機能検出あり", n_with_feat)
    c3.metric("スコア3以上（高複雑度）", n_high)
    c4.metric("最高スコア", int(df_app["アプリ性スコア"].max()) if len(df_app) else 0)

    # ── カテゴリ別ヒット数 ─────────────────────────────────────────────
    st.subheader("カテゴリ別検出ページ数")
    df_cat = pd.DataFrame(
        [{"カテゴリ": c, "検出ページ数": cat_counts[c]} for c in CATEGORIES if cat_counts[c] > 0]
    ).sort_values("検出ページ数", ascending=False)
    st.bar_chart(df_cat.set_index("カテゴリ")["検出ページ数"])

    # ── フィルター付きページ一覧 ───────────────────────────────────────
    st.subheader("ページ一覧（アプリ性スコア順）")
    _cat_opts = ["すべて"] + [c for c in CATEGORIES if cat_counts.get(c, 0) > 0]
    _cat_sel = st.selectbox("カテゴリでフィルタ", _cat_opts)
    _score_min = st.slider("アプリ性スコア（最小）", 0, len(CATEGORIES), 1)

    df_disp = df_app.copy()
    if _cat_sel != "すべて":
        df_disp = df_disp[df_disp[_cat_sel] != ""]
    df_disp = df_disp[df_disp["アプリ性スコア"] >= _score_min]
    df_disp = df_disp.sort_values("アプリ性スコア", ascending=False)

    # 表示用: 全カラムを表示するとワイドになるので要約カラムを優先
    _display_cols = ["URL", "タイトル", "アプリ性スコア", "検出機能"]
    st.dataframe(df_disp[_display_cols], use_container_width=True, height=460)

    # ── カテゴリ別ピボット（Excel向け） ────────────────────────────────
    with st.expander("カテゴリ別詳細（ピボット表示）"):
        _pivot_cols = ["URL", "タイトル", "アプリ性スコア"] + CATEGORIES
        st.dataframe(
            df_app.sort_values("アプリ性スコア", ascending=False)[_pivot_cols],
            use_container_width=True,
            height=400,
        )

    # ── ダウンロード ───────────────────────────────────────────────────
    _all_cols = ["URL", "タイトル", "アプリ性スコア", "検出機能"] + CATEGORIES
    df_dl = df_app.sort_values("アプリ性スコア", ascending=False)[_all_cols]
    cx, cy = st.columns(2)
    cx.download_button(
        "📥 Excel（全件）",
        to_excel_bytes(df_dl),
        "app_detection.xlsx",
        use_container_width=True,
        key="dl_app_xlsx",
    )
    cy.download_button(
        "📥 CSV（全件）",
        df_dl.to_csv(index=False).encode("utf-8-sig"),
        "app_detection.csv",
        mime="text/csv",
        use_container_width=True,
        key="dl_app_csv",
    )
