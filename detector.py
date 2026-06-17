"""Webアプリ・サービス機能検出モジュール"""
from __future__ import annotations
import re
from typing import Callable

# (タグ, カテゴリ, check_fn(url, html) -> bool)
DETECTORS: list[tuple[str, str, Callable[[str, str], bool]]] = []

# ── DSL ──────────────────────────────────────────────────────────────────────

def _h(*pats: str) -> Callable[[str, str], bool]:
    """HTML内にいずれかのパターンが存在する"""
    rs = [re.compile(p, re.I | re.S) for p in pats]
    return lambda _u, html: any(r.search(html) for r in rs)

def _u(*pats: str) -> Callable[[str, str], bool]:
    """URLがいずれかのパターンにマッチ"""
    rs = [re.compile(p, re.I) for p in pats]
    return lambda url, _h: any(r.search(url) for r in rs)

def _any(*checks: Callable) -> Callable[[str, str], bool]:
    return lambda url, html: any(c(url, html) for c in checks)

def _all(*checks: Callable) -> Callable[[str, str], bool]:
    return lambda url, html: all(c(url, html) for c in checks)

def _add(tag: str, cat: str, check: Callable) -> None:
    DETECTORS.append((tag, cat, check))


# ── JSフレームワーク ──────────────────────────────────────────────────────────
_add("Next.js",  "JSフレームワーク", _h(r"__NEXT_DATA__|/_next/static"))
_add("Nuxt.js",  "JSフレームワーク", _h(r"__NUXT__|/_nuxt/"))
_add("React",    "JSフレームワーク", _h(
    r"data-reactroot|react-dom(?:\.min)?\.js|__reactFiber|__reactProps|react\.production"))
_add("Vue.js",   "JSフレームワーク", _h(
    r"__vue__|vue(?:\.min)?\.js|vue\.runtime|data-v-[0-9a-f]{7,8}"))
_add("Angular",  "JSフレームワーク", _h(
    r'ng-version=|@angular/core|angular(?:\.min)?\.js'))

# ── 認証 ─────────────────────────────────────────────────────────────────────
_add("ログインページ", "認証", _any(
    _u(r"login|signin|sign.in|/auth"),
    _all(_h(r'type=["\']password'), _h(r"<form")),
))
_add("マイページ/会員", "認証", _any(
    _u(r"mypage|my.page|my.toyota|my-toyota|member|myaccount"),
    _h(r"マイページ|マイトヨタ|ログアウト|会員専用"),
))
_add("トヨタID/認証連携", "認証", _h(
    r"toyota-id|toyota_id|ToyotaID|id\.toyota|connect\.toyota"))

# ── フォーム ─────────────────────────────────────────────────────────────────
_add("問い合わせフォーム", "フォーム", _any(
    _u(r"contact|inquiry|inquire|toiawase"),
    _all(_h(r"<form"), _h(r"お問い合わせ|問い合わせ|Contact")),
))
_add("試乗・来店予約", "フォーム", _any(
    _u(r"testdrive|test.drive|showroom|reservation|reserve"),
    _h(r"試乗予約|来店予約|test\s*drive"),
))
_add("申込・資料請求", "フォーム", _any(
    _u(r"/apply|/form|/request|/booking"),
    _all(_h(r"<form"), _h(r"申込|申請|資料請求|見積もり依頼")),
))
_add("検索フォーム", "フォーム", _any(
    _h(r'type=["\']search["\']|role=["\']search["\']|<search[\s>]'),
    _u(r"/search[/?]"),
))

# ── 設定ツール・コンフィギュレーター ─────────────────────────────────────────
_add("グレード/カラー選択", "設定ツール", _any(
    _u(r"grade|colorselect|customize|simulator|color\.html"),
    _h(r"グレード選択|グレード・カラー|カラーシミュレーター|grade.+select"),
))
_add("見積もりシミュレーター", "設定ツール", _any(
    _u(r"estimate|simulation|simulator|loan|payment"),
    _h(r"見積もりシミュレーター|月々の支払|ローン計算|残価設定型"),
))
_add("販売店検索", "設定ツール", _any(
    _u(r"dealer|salesshop|storesearch|shopSearch"),
    _h(r"販売店を探す|ディーラー検索|店舗検索|nearest.{0,20}dealer"),
))
_add("中古車・在庫検索", "設定ツール", _any(
    _u(r"used|usedcar|cpo|certified|chuko"),
    _h(r"中古車|認定中古|CPO|在庫を探す|在庫検索"),
))
_add("オーナーズ/メンテナンス", "設定ツール", _any(
    _u(r"owner|maintenance|parts|service.schedule"),
    _h(r"オーナーズ|車検|点検予約|メンテナンスパック"),
))

# ── メディア ─────────────────────────────────────────────────────────────────
_add("動画（Brightcove）", "メディア", _h(
    r"brightcove\.net|players\.brightcove|videocloud|bcove\.me"))
_add("動画（YouTube/Vimeo）", "メディア", _h(
    r"youtube\.com/embed|youtu\.be|player\.vimeo\.com"))
_add("動画（HTML5）", "メディア", _h(r"<video[\s>]"))
_add("インタラクティブマップ", "メディア", _h(
    r"maps\.googleapis\.com|google\.com/maps|mapbox|leaflet\.js"))
_add("360°ビュー/3D", "メディア", _h(
    r"360|three\.js|webgl|panorama|spincar|evox"))

# ── SPA/PWA ──────────────────────────────────────────────────────────────────
_add("PWA対応", "SPA/PWA", _any(
    _h(r'rel=["\']manifest["\']'),
    _h(r"serviceWorker\.register|navigator\.serviceWorker"),
))
_add("SPAルーティング", "SPA/PWA", _h(
    r"history\.pushState|history\.replaceState|vue-router|@angular/router"))
_add("WebAPI/Fetch多用", "SPA/PWA", _any(
    _h(r"fetch\(['\"]https?://"),
    _h(r"XMLHttpRequest|axios\.(?:get|post)|apollo.+client"),
))

# ── 外部サービス ──────────────────────────────────────────────────────────────
_add("チャット/チャットボット", "外部サービス", _h(
    r"zendesk|intercom|livechat|freshchat|chatbot|チャットボット|line.*chat\.js"))
_add("SNS共有ボタン", "外部サービス", _h(
    r"twitter\.com/(?:share|intent)|facebook\.com/sharer|line\.me/R/msg"))
_add("外部iframe埋め込み", "外部サービス", _h(
    r'<iframe[^>]+src=["\']https?://(?!(?:www\.)?toyota\.jp)'))
_add("GTM/Adobe Analytics", "外部サービス", _any(
    _h(r"googletagmanager\.com/gtm\.js|gtag\("),
    _h(r"adobe.*launch|AppMeasurement\.js|s\.t\(\)"),
))
_add("電話クリックtoコール", "外部サービス", _h(r'href=["\']tel:'))

# ── トヨタ固有 ────────────────────────────────────────────────────────────────
_add("カーライン選択UI", "トヨタ固有", _any(
    _u(r"/(?:all|lineup|carlineup)"),
    _h(r"カーラインアップ|全車種一覧|carlineup"),
))
_add("T-Connect/コネクテッド", "トヨタ固有", _h(
    r"t-connect|tconnect|コネクテッド|DCM|G-BOOK"))
_add("サブスクリプション（KINTO）", "トヨタ固有", _h(
    r"kinto|キント|サブスク|月額|定額"))
_add("ファイナンス/ローン", "トヨタ固有", _any(
    _u(r"finance|financing|loan"),
    _h(r"トヨタファイナンス|ローン|リース|頭金"),
))


# ── Public API ────────────────────────────────────────────────────────────────

CATEGORIES: list[str] = list(dict.fromkeys(cat for _, cat, _ in DETECTORS))


def detect(url: str, html: str) -> list[str]:
    """検出されたタグのリストを返す"""
    return [tag for tag, _, chk in DETECTORS if chk(url, html)]


def detect_with_category(url: str, html: str) -> list[dict]:
    """[{"タグ": ..., "カテゴリ": ...}, ...] の形式で返す"""
    return [{"タグ": tag, "カテゴリ": cat}
            for tag, cat, chk in DETECTORS if chk(url, html)]


def summarize(url: str, html: str, title: str = "") -> dict:
    """1ページの検出結果を要約辞書で返す（DataFrameの1行分）"""
    by_cat: dict[str, list[str]] = {c: [] for c in CATEGORIES}
    for tag, cat, chk in DETECTORS:
        if chk(url, html):
            by_cat[cat].append(tag)

    detected_cats = [c for c in CATEGORIES if by_cat[c]]
    row: dict = {
        "URL": url,
        "タイトル": title,
        "アプリ性スコア": len(detected_cats),
        "検出機能": " / ".join(
            f"[{c}] {', '.join(by_cat[c])}" for c in CATEGORIES if by_cat[c]
        ),
    }
    for cat in CATEGORIES:
        row[cat] = ", ".join(by_cat[cat]) if by_cat[cat] else ""
    return row
