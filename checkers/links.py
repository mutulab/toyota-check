"""リンク切れチェッカー"""

import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse
from fetcher import fetch_html, extract_links
from config import CONCURRENT_WORKERS, REQUEST_DELAY


def check_links(urls: list[str], deep: bool = False) -> list[dict]:
    """
    URLリストのリンク切れをチェック。
    deep=True の場合、各ページ内のリンクも検査。
    """
    results = []
    checked = set()

    def _check_one(url: str, source: str = "") -> dict:
        status, html = fetch_html(url)
        result = {
            "url": url,
            "source": source,
            "status": status,
            "ok": status in (200, 301, 302),
            "broken": status in (404, 410) or status == 0,
            "title": "",
        }
        return result, html

    # 1st pass: 入力URLを検査
    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futures = {ex.submit(_check_one, url, "入力リスト"): url for url in urls}
        for future in as_completed(futures):
            result, html = future.result()
            results.append(result)
            checked.add(result["url"])
            time.sleep(REQUEST_DELAY)

    # deep=Trueのときはページ内リンクも検査
    if deep:
        inner_links = set()
        for r in results:
            if r["ok"] and r.get("html"):
                for link in extract_links(r["html"], r["url"]):
                    if urlparse(link).netloc == "toyota.jp" and link not in checked:
                        inner_links.add(link)

        with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
            futures = {
                ex.submit(_check_one, url, "ページ内リンク"): url
                for url in inner_links
            }
            for future in as_completed(futures):
                result, _ = future.result()
                results.append(result)
                time.sleep(REQUEST_DELAY)

    broken = [r for r in results if r["broken"]]
    print(f"  チェック済み: {len(results)}件 / リンク切れ: {len(broken)}件")
    return results
