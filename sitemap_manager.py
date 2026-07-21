"""サイトマップ管理モード

tjpコンテンツ管理表.xlsx（運用サイトマップ）をアプリ上で閲覧・編集・保守する。
  - 第6階層までの階層列を自動生成し、表示階層を指定してビューを切替
  - URL一覧（編集可・行追加可）／ディレクトリ集計の2ビュー
  - CSV / Excel ダウンロード
  - HTML（静的）かアプリケーション（動的）かをURLルールで自動判別
  - サイトをクロールして一覧に無いURL・より深い階層を検知し、一覧へ追加
"""

from __future__ import annotations

import io
import re
import time
from collections import deque
from datetime import date
from urllib.parse import urlparse, urljoin, urldefrag

import pandas as pd

SHEET_NAME = "運用サイトマップ"
HEADER_ROW = 5          # 見出し行（1始まり）
DATA_MIN_ROW = 6        # データ開始行
URL_COL_NAME = "フルURL"
MAX_LEVEL = 6

# TID移行対象_サービスのまとまり整理 v1.0 突合済みのアプリ入口パス
APP_PREFIXES = [
    "/service/", "/member/", "/login", "/profile", "/grade",
    "/cmpn", "/socialfes", "/follow", "/ucar_search", "/webservice",
    "/mailalert_service", "/measurement", "/faq/inquiry",
]
APP_EXACT = {"/recall", "/recall/"}   # /recall はトップのみアプリ（配下は静的届出ページ）

FILE_EXTS = (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".zip",
             ".xlsx", ".xls", ".docx", ".doc", ".pptx", ".mp4", ".webp", ".ico")


# ─── 純粋ロジック（テスト可能） ──────────────────────────────

def classify_url(url: str) -> str:
    """URLから HTML（静的） / アプリ（動的） / ファイル を判別する。"""
    try:
        p = urlparse(str(url))
    except Exception:
        return "不明"
    path = p.path or "/"
    low = path.lower()
    if p.netloc and "toyota.jp" not in p.netloc and "toyota-catalog.jp" not in p.netloc:
        return "外部サイト"
    if low.rstrip("/") in APP_EXACT or low in APP_EXACT:
        return "アプリ"
    for pref in APP_PREFIXES:
        if low == pref or low.startswith(pref if pref.endswith("/") else pref + "/") \
           or low.rstrip("/") == pref.rstrip("/"):
            return "アプリ"
    if low.endswith(FILE_EXTS):
        return "ファイル"
    if p.query and not low.endswith((".html", ".htm")):
        return "アプリ"          # クエリ付き動的URL（dc/search 等）
    return "HTML（静的）"


def path_levels(url: str, n: int = MAX_LEVEL) -> list[str]:
    """URLのパスを第1〜第n階層に分解して返す（不足分は空文字）。"""
    try:
        p = urlparse(str(url))
    except Exception:
        return [""] * n
    segs = [s for s in (p.path or "").split("/") if s]
    if segs and segs[-1] in ("index.html", "index.htm"):
        segs = segs[:-1]
    segs = segs[:n]
    return segs + [""] * (n - len(segs))


def norm_url(u: str) -> str:
    """比較用の正規化（フラグメント除去・index.html除去・末尾スラッシュ無視）。"""
    u, _ = urldefrag(str(u).strip())
    u = re.sub(r"/index\.html?$", "/", u)
    u = re.sub(r"^http://", "https://", u)
    return u.rstrip("/").lower()


def load_sitemap(file) -> tuple[pd.DataFrame, list[str]]:
    """運用サイトマップシートを全列DataFrameとして読み込む。"""
    import openpyxl
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"シート「{SHEET_NAME}」が見つかりません")
    ws = wb[SHEET_NAME]
    rows = list(ws.iter_rows(min_row=HEADER_ROW, values_only=True))
    raw_headers = list(rows[0])
    headers: list[str] = []
    for i, h in enumerate(raw_headers):
        name = str(h).strip() if h not in (None, "") else f"列{i + 1}"
        while name in headers:
            name += "_"
        headers.append(name)
    data = [list(r) + [None] * (len(headers) - len(r)) for r in rows[1:]]
    df = pd.DataFrame(data, columns=headers)
    if URL_COL_NAME not in df.columns:
        raise ValueError(f"「{URL_COL_NAME}」列が見つかりません")
    # URLが空の行は除外（相対パスは補完）
    df[URL_COL_NAME] = df[URL_COL_NAME].astype("string").str.strip()
    df = df[df[URL_COL_NAME].notna() & (df[URL_COL_NAME] != "")].reset_index(drop=True)
    df[URL_COL_NAME] = df[URL_COL_NAME].map(
        lambda u: "https://toyota.jp" + u if str(u).startswith("/") else u)
    if "転記元" not in "".join(df.columns):
        df["転記元"] = None
    return df, headers


def add_derived(df: pd.DataFrame) -> pd.DataFrame:
    """階層1〜6・階層深さ・種別判定の派生列を付与する。"""
    out = df.copy()
    levels = out[URL_COL_NAME].map(lambda u: path_levels(u))
    for i in range(MAX_LEVEL):
        out[f"階層{i + 1}"] = levels.map(lambda L, i=i: L[i])
    out["階層深さ"] = levels.map(lambda L: sum(1 for s in L if s))
    out["種別判定"] = out[URL_COL_NAME].map(classify_url)
    return out


def agg_directories(dfd: pd.DataFrame, level: int) -> pd.DataFrame:
    """指定階層までのディレクトリ単位に集計する。"""
    keys = [f"階層{i + 1}" for i in range(level)]
    g = dfd.groupby(keys, dropna=False)
    rows = []
    for name, grp in g:
        name = name if isinstance(name, tuple) else (name,)
        path = "/" + "/".join(s for s in name if s)
        if path != "/":
            path += "/"
        rows.append({
            "ディレクトリ": path,
            "ページ数": len(grp),
            "HTML": int((grp["種別判定"] == "HTML（静的）").sum()),
            "アプリ": int((grp["種別判定"] == "アプリ").sum()),
            "その他": int((~grp["種別判定"].isin(["HTML（静的）", "アプリ"])).sum()),
            "代表ページ": str(grp.iloc[0].get("ページ説明") or "")[:40],
        })
    return (pd.DataFrame(rows)
            .sort_values("ページ数", ascending=False)
            .reset_index(drop=True))


def crawl_discover(existing_norm: set[str], start_url: str, path_filter: str,
                   max_pages: int, max_depth: int,
                   progress_cb=None) -> list[dict]:
    """クロールして一覧に無いURLを検知する。toyota.jp内のみ・HTMLのみ巡回。"""
    from fetcher import fetch_html, extract_links, extract_meta
    from config import REQUEST_DELAY

    seen: set[str] = set()
    found: dict[str, dict] = {}
    q: deque[tuple[str, int]] = deque([(start_url, 0)])
    pages = 0
    while q and pages < max_pages:
        url, depth = q.popleft()
        nu = norm_url(url)
        if nu in seen:
            continue
        seen.add(nu)
        try:
            status, html = fetch_html(url)
        except Exception:
            continue
        pages += 1
        if progress_cb:
            progress_cb(pages, max_pages, url, len(found))
        if status != 200 or not html:
            continue
        title = ""
        try:
            title = (extract_meta(html).get("title") or "").strip()
        except Exception:
            pass
        if nu not in existing_norm and nu != norm_url(start_url):
            found[nu] = {"フルURL": url.split("#")[0], "ページ説明": title,
                         "階層深さ": sum(1 for s in path_levels(url) if s),
                         "種別判定": classify_url(url), "リンク元深さ": depth}
        if depth >= max_depth:
            continue
        try:
            links = extract_links(html, url)
        except Exception:
            continue
        for link in links:
            lp = urlparse(link)
            if lp.netloc not in ("toyota.jp", "www.toyota.jp"):
                continue
            if path_filter and not lp.path.startswith(path_filter):
                continue
            if lp.path.lower().endswith(FILE_EXTS):
                continue
            nl = norm_url(link)
            if nl not in seen:
                q.append((link, depth + 1))
        time.sleep(REQUEST_DELAY)
    return sorted(found.values(), key=lambda d: d["フルURL"])


def build_tree(dfd: pd.DataFrame, level: int) -> pd.DataFrame:
    """パス順に並べ、インデント付きツリー列と配下ページ数を持つ表を作る。

    level より深い行は畳む（親の「配下ページ数」に含まれるため情報は失われない）。
    """
    d = dfd.copy()
    d["_segs"] = d[URL_COL_NAME].map(lambda u: tuple(s for s in path_levels(u) if s))
    d = d.sort_values("_segs", kind="stable").reset_index(drop=True)

    # 配下ページ数: 自分のパスを接頭辞とする行数（自分を除く）
    seg_list = list(d["_segs"])
    counts = []
    for segs in seg_list:
        n = sum(1 for o in seg_list if len(o) > len(segs) and o[:len(segs)] == segs)
        counts.append(n)
    d["配下ページ数"] = counts

    shown = d[d["階層深さ"] <= level].copy()

    def tree_label(row):
        segs = row["_segs"]
        depth = len(segs)
        if depth == 0:
            return "🏠 toyota.jp（トップ）"
        name = segs[-1]
        url = str(row[URL_COL_NAME])
        is_dir = url.rstrip("/").endswith(name) and url.endswith("/")
        icon = "📁 " if (is_dir or row["配下ページ数"] > 0) else "📄 "
        return "　　" * (depth - 1) + "└ " + icon + name + ("/" if is_dir else "")

    shown["ツリー"] = shown.apply(tree_label, axis=1)
    folded = shown["配下ページ数"].where(shown["階層深さ"] == level, 0)
    shown["畳まれた配下"] = folded.map(lambda n: f"+{n}件" if n else "")
    cols = ["ツリー", "配下ページ数", "畳まれた配下", "種別判定", "ページ説明", URL_COL_NAME]
    return shown[[c for c in cols if c in shown.columns]]


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=SHEET_NAME)
    return buf.getvalue()


# ─── Streamlit UI ────────────────────────────────────────────

VIEW_COLS = ["№", URL_COL_NAME, "ページ説明", "種別判定", "階層深さ",
             "運用対象コンテンツ", "更新頻度", "情報主幹部署", "運用会社", "アプリ"]


def render():
    import streamlit as st

    st.header("🗂️ サイトマップ管理")
    st.caption("tjpコンテンツ管理表（運用サイトマップ）の閲覧・編集・クロール差分検知")

    # ── 1. 読み込み ──
    up = st.file_uploader("tjpコンテンツ管理表.xlsx をアップロード", type=["xlsx"],
                          key="smgr_upload")
    if up is not None:
        fkey = f"{up.name}_{up.size}"
        if st.session_state.get("smgr_fkey") != fkey:
            try:
                df, headers = load_sitemap(up)
            except Exception as e:
                st.error(f"読み込みエラー: {e}")
                return
            st.session_state["smgr_df"] = df
            st.session_state["smgr_fkey"] = fkey
            st.session_state.pop("smgr_new_urls", None)

    df: pd.DataFrame | None = st.session_state.get("smgr_df")
    if df is None:
        st.info("コンテンツ管理表（.xlsx）をアップロードしてください。"
                "編集内容はセッション中保持され、CSV / Excel でダウンロードできます。")
        return

    dfd = add_derived(df)

    # ── 2. サマリー ──
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("URL数", len(dfd))
    c2.metric("HTML（静的）", int((dfd["種別判定"] == "HTML（静的）").sum()))
    c3.metric("アプリ", int((dfd["種別判定"] == "アプリ").sum()))
    c4.metric("最大階層", int(dfd["階層深さ"].max() or 0))
    c5.metric("第6階層以深", int((dfd["階層深さ"] >= 6).sum()))

    # ── 3. ビュー設定 ──
    st.divider()
    f1, f2, f3, f4 = st.columns([1.4, 1.2, 1.2, 1.4])
    view_mode = f1.radio("表示モード",
                         ["🌲 ツリービュー", "✏️ 一覧編集", "📊 ディレクトリ集計"],
                         key="smgr_mode")
    level = f2.select_slider("表示階層",
                             options=list(range(1, MAX_LEVEL + 1)), value=MAX_LEVEL,
                             format_func=lambda n: f"第{n}階層まで",
                             key="smgr_level")
    dir1 = f3.selectbox(
        "第1階層で絞り込み",
        ["（すべて）"] + sorted(x for x in dfd["階層1"].unique() if x),
        key="smgr_dir1")
    q = f4.text_input("URL・説明で検索", key="smgr_q").strip()

    # 絞り込みは「検索・第1階層」のみ。階層スライダーで行が消えるのはツリーの折り畳みだけ
    view = dfd
    if dir1 != "（すべて）":
        view = view[view["階層1"] == dir1]
    if q:
        mask = (view[URL_COL_NAME].astype(str).str.contains(q, case=False, na=False)
                | view.get("ページ説明", pd.Series("", index=view.index))
                    .astype(str).str.contains(q, case=False, na=False))
        view = view[mask]

    filtered = len(view) != len(dfd)
    fc1, fc2 = st.columns([4, 1])
    fc1.caption(f"全 {len(dfd)} 件中 **{len(view)} 件** を表示"
                + ("（絞り込み中）" if filtered else ""))
    if filtered and fc2.button("絞り込み解除", key="smgr_reset"):
        st.session_state["smgr_dir1"] = "（すべて）"
        st.session_state["smgr_q"] = ""
        st.rerun()

    # ── 4. 表示・編集 ──
    if view_mode == "📊 ディレクトリ集計":
        st.subheader(f"ディレクトリ集計（第{level}階層まで）")
        agg = agg_directories(view, level)
        st.dataframe(agg, use_container_width=True, height=480)
        st.download_button("⬇️ この集計をCSVダウンロード", to_csv_bytes(agg),
                           f"sitemap_dir_level{level}_{date.today()}.csv", "text/csv")
    elif view_mode == "🌲 ツリービュー":
        tree = build_tree(view, level)
        hidden = len(view) - len(tree)
        st.subheader(f"ツリービュー（第{level}階層まで表示・{len(tree)} 行）")
        if hidden:
            st.caption(f"※ 第{level + 1}階層より深い {hidden} 件は畳んでいます"
                       "（「畳まれた配下」列に件数表示。表示階層を上げると展開）")
        st.dataframe(
            tree, use_container_width=True, height=520, hide_index=True,
            column_config={
                "ツリー": st.column_config.TextColumn("サイト構造", width="large"),
                "配下ページ数": st.column_config.NumberColumn("配下", width="small"),
                "畳まれた配下": st.column_config.TextColumn("畳み", width="small"),
                "種別判定": st.column_config.TextColumn("種別", width="small"),
                URL_COL_NAME: st.column_config.LinkColumn("URL", width="medium"),
            },
        )
        st.download_button("⬇️ このツリーをCSV", to_csv_bytes(tree),
                           f"sitemap_tree_level{level}_{date.today()}.csv", "text/csv",
                           key="smgr_tree_csv")
    else:
        show_all = st.checkbox("全列を表示（Excelの全項目）", key="smgr_allcols")
        lvl_cols = [f"階層{i + 1}" for i in range(level)]
        cols = ([c for c in VIEW_COLS if c in view.columns] + lvl_cols) if not show_all \
            else list(view.columns)
        st.subheader(f"一覧編集（{len(view)} 件）")
        edited = st.data_editor(
            view[cols], use_container_width=True, height=480,
            num_rows="dynamic", key="smgr_editor",
            disabled=["種別判定", "階層深さ"] + lvl_cols,
        )
        if st.button("💾 編集内容を保存（セッションに反映）", type="primary"):
            base = df.copy()
            editable = [c for c in edited.columns
                        if c in base.columns and c not in ("種別判定", "階層深さ")]
            common = edited.index.intersection(base.index)
            base.loc[common, editable] = edited.loc[common, editable]
            new_rows = edited.loc[edited.index.difference(base.index), editable]
            new_rows = new_rows[new_rows[URL_COL_NAME].notna()]
            if len(new_rows):
                base = pd.concat([base, new_rows], ignore_index=True)
            st.session_state["smgr_df"] = base
            st.success(f"保存しました（更新 {len(common)} 行・追加 {len(new_rows)} 行）。"
                       "ダウンロードで書き出せます。")
            st.rerun()

    # ── 5. ダウンロード ──
    d1, d2, d3 = st.columns(3)
    d1.download_button("⬇️ 表示中の絞り込み結果をCSV", to_csv_bytes(view),
                       f"sitemap_view_{date.today()}.csv", "text/csv")
    d2.download_button("⬇️ 全件CSV（判定・階層列つき）", to_csv_bytes(add_derived(df)),
                       f"sitemap_all_{date.today()}.csv", "text/csv")
    d3.download_button("⬇️ Excel（編集反映済み）", to_excel_bytes(df),
                       f"tjpコンテンツ管理表_更新_{date.today()}.xlsx")

    # ── 6. クロール差分検知 ──
    st.divider()
    with st.expander("🕷️ クロールして一覧に無いURLを検知・追加", expanded=False):
        cc1, cc2, cc3, cc4 = st.columns([1.6, 1.2, 1, 1])
        start_url = cc1.text_input("開始URL", "https://toyota.jp/", key="smgr_c_start")
        path_filter = cc2.text_input("パス制限（任意）", "", placeholder="/carlineup/",
                                     key="smgr_c_filter").strip()
        max_pages = cc3.number_input("最大巡回ページ数", 10, 2000, 200, step=10,
                                     key="smgr_c_max")
        max_depth = cc4.number_input("リンク深さ", 1, MAX_LEVEL, 3, key="smgr_c_depth")
        st.caption("toyota.jp内のHTMLのみ巡回します。一覧との差分（未登録URL）を検知します。")

        if st.button("▶ クロール開始", key="smgr_c_run"):
            existing = set(dfd[URL_COL_NAME].map(norm_url))
            bar = st.progress(0.0)
            info = st.empty()

            def cb(done, total, url, nfound):
                bar.progress(min(done / total, 1.0))
                info.caption(f"{done}/{total} ページ巡回中… 未登録検知 {nfound} 件　{url[:80]}")

            new_urls = crawl_discover(existing, start_url, path_filter,
                                      int(max_pages), int(max_depth), cb)
            st.session_state["smgr_new_urls"] = new_urls
            bar.progress(1.0)

        new_urls = st.session_state.get("smgr_new_urls")
        if new_urls is not None:
            if not new_urls:
                st.success("一覧に無いURLは見つかりませんでした。")
            else:
                st.warning(f"一覧に無いURLを {len(new_urls)} 件検知しました。"
                           "追加する行にチェックを入れてください。")
                nd = pd.DataFrame(new_urls)
                nd.insert(0, "追加", True)
                sel = st.data_editor(nd, use_container_width=True, height=320,
                                     key="smgr_new_editor",
                                     disabled=[c for c in nd.columns if c != "追加"])
                a1, a2 = st.columns([1, 2])
                if a1.button("➕ チェックした行を一覧に追加", type="primary",
                             key="smgr_add_btn"):
                    add = sel[sel["追加"]]
                    base = df.copy()
                    src_col = next((c for c in base.columns if "転記元" in c), None)
                    rows = []
                    for _, r in add.iterrows():
                        row = {c: None for c in base.columns}
                        row[URL_COL_NAME] = r["フルURL"]
                        if "ページ説明" in base.columns:
                            row["ページ説明"] = r["ページ説明"]
                        if src_col:
                            row[src_col] = f"クロール検知 {date.today()}"
                        rows.append(row)
                    base = pd.concat([base, pd.DataFrame(rows)], ignore_index=True)
                    st.session_state["smgr_df"] = base
                    st.session_state["smgr_new_urls"] = [
                        u for u in new_urls
                        if u["フルURL"] not in set(add["フルURL"])]
                    st.success(f"{len(rows)} 件を一覧に追加しました。"
                               "ダウンロードでExcel/CSVに書き出せます。")
                    st.rerun()
                a2.download_button("⬇️ 検知結果をCSV", to_csv_bytes(nd.drop(columns=['追加'])),
                                   f"crawl_new_urls_{date.today()}.csv", "text/csv",
                                   key="smgr_new_csv")
