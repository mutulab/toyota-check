"""toyota-check 設定"""

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "identity",
    "Referer": "https://toyota.jp/",
}

REQUEST_TIMEOUT = 15       # 秒
CONCURRENT_WORKERS = 5     # 同時リクエスト数
REQUEST_DELAY = 0.3        # リクエスト間隔（秒）

# PSI API（Core Web Vitals）
PSI_API_KEY = ""           # 空欄でも動作（上限25,000回/日→要APIキー）
PSI_STRATEGY = "mobile"    # "mobile" or "desktop"
PSI_ENDPOINT = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"

# TQP KPI 閾値
THRESHOLDS = {
    "LCP":  {"good": 2500,  "poor": 4000},
    "CLS":  {"good": 0.1,   "poor": 0.25},
    "INP":  {"good": 200,   "poor": 500},
    "FCP":  {"good": 1800,  "poor": 3000},
    "TTFB": {"good": 800,   "poor": 1800},
}

# 禁止表現パターン
KINSHI_PATTERNS = [
    (r"日本一|世界一|No\.?1|ナンバーワン", "優位性表現（要根拠確認）"),
    (r"必ず|絶対に|100%", "断定的表現（要確認）"),
    (r"最高|最大|最速|最安", "最上級表現（要根拠確認）"),
    (r"無料|タダ|0円", "無料表現（要確認）"),
    (r"期間限定|今だけ|本日限り", "限定表現（要確認）"),
]

# 表記ゆれチェック辞書（正規表現: 推奨表記）
HYOKI_YURE = {
    r"Webサイト|web\s*サイト|ウェブサイト": "WEBサイト",
    r"E-?mail|e-?mail|メール": "メール",
    r"お問い合わせ|お問合せ|お問合わせ": "お問い合わせ",
    r"ログイン|ろぐいん|log\s*in": "ログイン",
    r"ホームページ": "ホームページ（※WEBサイトに統一検討）",
}

# リンクチェック除外パターン
LINK_IGNORE_PATTERNS = [
    r"^mailto:",
    r"^tel:",
    r"^javascript:",
    r"^#",
]
