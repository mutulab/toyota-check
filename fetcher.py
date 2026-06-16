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
    """ページ内の<a href>リンクを絶対URLで返す"""
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
    return list(dict.fromkeys(links))


def extract_resources(html: str, base_url: str) -> list[tuple[str, str]]:
    """ページ内の全リソース（リンク・CSS・JS・画像・メディア）を (url, 種別) で返す"""
    from urllib.parse import urljoin, urlparse

    SKIP = ("mailto:", "tel:", "javascript:", "data:")
    seen: set[str] = set()
    results: list[tuple[str, str]] = []

    def _add(href: str, rtype: str) -> None:
        if not href or any(href.startswith(p) for p in SKIP):
            return
        abs_url = urljoin(base_url, href.split("#")[0])
        if urlparse(abs_url).scheme not in ("http", "https"):
            return
        if abs_url not in seen:
            seen.add(abs_url)
            results.append((abs_url, rtype))

    patterns = [
        (r'<a[^>]+href="([^"]*)"',              "リンク"),
        (r'<link[^>]+href="([^"]*)"',            "CSS/スタイル"),
        (r'<script[^>]+src="([^"]*)"',           "JavaScript"),
        (r'<img[^>]+src="([^"]*)"',              "画像"),
        (r'<source[^>]+src="([^"]*)"',           "メディア"),
        (r'<(?:video|audio)[^>]+src="([^"]*)"',  "メディア"),
        (r'<iframe[^>]+src="([^"]*)"',           "iframe"),
    ]
    for pattern, rtype in patterns:
        for m in re.finditer(pattern, html, re.I):
            _add(m.group(1), rtype)

    # インラインstyle属性内のurl()
    for m in re.finditer(r'style="[^"]*url\(["\']?([^"\')\s]+)["\']?\)', html, re.I):
        _add(m.group(1), "画像(インラインCSS)")

    return results


def extract_css_urls(css_text: str, base_url: str) -> list[tuple[str, str]]:
    """CSSテキスト内のurl()参照を (url, 種別) で返す"""
    from urllib.parse import urljoin, urlparse

    SKIP = ("data:", "#")
    seen: set[str] = set()
    results: list[tuple[str, str]] = []

    for m in re.finditer(r'url\(\s*["\']?([^"\')\s]+)["\']?\s*\)', css_text, re.I):
        href = m.group(1).strip()
        if any(href.startswith(p) for p in SKIP):
            continue
        abs_url = urljoin(base_url, href)
        parsed = urlparse(abs_url)
        if parsed.scheme not in ("http", "https"):
            continue
        if abs_url in seen:
            continue
        seen.add(abs_url)
        path = parsed.path.lower()
        if any(path.endswith(ext) for ext in (".woff", ".woff2", ".ttf", ".eot", ".otf")):
            rtype = "フォント(CSS)"
        elif any(path.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".avif")):
            rtype = "画像(CSS)"
        else:
            rtype = "CSSリソース"
        results.append((abs_url, rtype))

    return results


def extract_text_blocks(html: str) -> list[str]:
    """本文テキストブロックを抽出（script/style除去）"""
    html = re.sub(r"<script[^>]*>.*?</script>", "", html, flags=re.S | re.I)
    html = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.S | re.I)
    html = re.sub(r"<[^>]+>", " ", html)
    html = re.sub(r"\s+", " ", html)
    return [b.strip() for b in html.split("。") if len(b.strip()) > 5]
