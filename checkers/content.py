"""コンテンツ検証（表記ゆれ・禁止表現・title/description取得）"""

import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from fetcher import fetch_html, extract_meta, extract_text_blocks
from config import HYOKI_YURE, CONCURRENT_WORKERS, REQUEST_DELAY

# 景表法・誇大表現チェックパターン（toyota.jp 運用ルール）
KINSHI_PATTERNS = [
    (r"日本一|世界一|No\.?1|ナンバーワン", "優位性表現（要根拠確認）"),
    (r"必ず|絶対に|100%", "断定的表現（要確認）"),
    (r"最高|最大|最速|最安", "最上級表現（要根拠確認）"),
    (r"無料|タダ|0円", "無料表現（要確認）"),
    (r"期間限定|今だけ|本日限り", "限定表現（要確認）"),
]


def check_content(urls: list[str]) -> list[dict]:
    """title・description取得 + 表記ゆれ・禁止表現チェック"""
    results = []

    def _process(url: str) -> dict:
        status, html = fetch_html(url)
        if status != 200 or not html:
            return {"url": url, "status": status, "error": True}

        meta = extract_meta(html)
        text_blocks = extract_text_blocks(html)
        full_text = " ".join(text_blocks)

        # 表記ゆれチェック
        yure_found = []
        for pattern, correct in HYOKI_YURE.items():
            matches = re.findall(pattern, full_text, re.I)
            if matches:
                yure_found.append({
                    "pattern": pattern,
                    "correct": correct,
                    "matches": list(set(matches))[:3],
                })

        # 禁止表現チェック
        kinshi_found = []
        for pattern, reason in KINSHI_PATTERNS:
            matches = re.findall(pattern, full_text)
            if matches:
                kinshi_found.append({
                    "pattern": pattern,
                    "reason": reason,
                    "matches": list(set(matches))[:3],
                })

        # title・description の問題チェック
        title_issues = []
        if not meta["short_title"]:
            title_issues.append("titleタグなし")
        elif len(meta["short_title"]) > 60:
            title_issues.append(f"title長すぎ({len(meta['short_title'])}文字)")
        if not meta["description"]:
            title_issues.append("descriptionなし")
        elif len(meta["description"]) > 160:
            title_issues.append(f"description長すぎ({len(meta['description'])}文字)")

        return {
            "url": url,
            "status": status,
            "title": meta["short_title"],
            "description": meta["description"][:100],
            "title_issues": title_issues,
            "hyoki_yure": yure_found,
            "kinshi_hyogen": kinshi_found,
            "ng": bool(yure_found or kinshi_found or title_issues),
        }

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futures = {ex.submit(_process, url): url for url in urls}
        done = 0
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            done += 1
            if done % 10 == 0:
                print(f"  進捗: {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)

    ng_count = sum(1 for r in results if r.get("ng"))
    print(f"  完了: {len(results)}件 / 要確認: {ng_count}件")
    return results


def fetch_titles_bulk(urls: list[str]) -> dict[str, dict]:
    """URLリストのtitle/descriptionを一括取得してdict返却"""
    results = {}

    def _fetch(url):
        status, html = fetch_html(url)
        if status == 200 and html:
            return url, extract_meta(html)
        return url, None

    with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as ex:
        futures = {ex.submit(_fetch, url): url for url in urls}
        done = 0
        for future in as_completed(futures):
            url, meta = future.result()
            results[url] = meta
            done += 1
            if done % 20 == 0:
                print(f"  取得中: {done}/{len(urls)}")
            time.sleep(REQUEST_DELAY)

    return results
