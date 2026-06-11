"""共通HTTPフェッチャー（ブラウザヘッダーでtoyota.jpの403を回避）"""

import re
import time
import urllib.request
import urllib.error
from config import HEADERS, REQUEST_TIMEOUT


def fetch_html(url: str) -> tuple[int, str]:
    """URLを取得してHTTPステータスとHTMLを返す"""
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as res:
            charset = _detect_charset(res.headers.get("Content-Type", ""))
            return res.status, res.read().decode(charset, errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, ""
    except Exception:
        return 0, ""


def _detect_charset(content_type: str) -> str:
    m = re.search(r"charset=([\w-]+)", content_type, re.I)
    return m.group(1) if m else "utf-8"


def extract_meta(html: str) -> dict:
    """title / description / h1 を抽出"""
    title = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
    desc = re.search(
        r'<meta[^>]+name="description"[^>]*content="([^"]+)"', html, re.I
    )
    if not desc:
        desc = re.search(
            r'<meta[^>]+content="([^"]+)"[^>]*name="description"', html, re.I
        )
    h1 = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.S | re.I)

    def clean(s):
        return re.sub(r"<[^>]+>", "", s).strip() if s else ""

    raw_title = clean(title.group(1)) if title else ""
    # "○○ | トヨタ自動車WEBサイト" → "○○" に整形
    short_title = re.sub(r"\s*[|｜]\s*トヨタ自動車WEBサイト.*$", "", raw_title).strip()

    return {
        "title": raw_title,
        "short_title": short_title,
        "description": clean(desc.group(1)) if desc else "",
        "h1": clean(h1.group(1)) if h1 else "",
    }


def extract_links(html: str, base_url: str) -> list[str]:
    """ページ内の全リンクを絶対URLで返す"""
    from urllib.parse import urljoin, urlparse
    import re as _re

    hrefs = _re.findall(r'<a[^>]+href="([^"#][^"]*)"', html, _re.I)
    links = []
    for href in hrefs:
        if href.startswith(("mailto:", "tel:", "javascript:")):
            continue
        abs_url = urljoin(base_url, href.split("#")[0])
        parsed = urlparse(abs_url)
        if parsed.scheme in ("http", "https"):
            links.append(abs_url)
    return list(dict.fromkeys(links))  # 重複除去


def extract_text_blocks(html: str) -> list[str]:
    """本文テキストブロックを抽出（script/style除去）"""
    html = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S | re.I)
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.S | re.I)
    html = re.sub(r"<[^>]+>", " ", html)
    html = re.sub(r"\s+", " ", html)
    return [b.strip() for b in html.split("。") if len(b.strip()) > 5]
