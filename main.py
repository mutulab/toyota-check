#!/usr/bin/env python3
"""
toyota-check — toyota.jp 専用サイト検証ツール
SPIRAL ISSO 相当機能をtoyota.jp向けに特化実装

使い方:
  python main.py titles              # Excelのページ説明をtitle/descで一括更新
  python main.py links               # リンク切れチェック（管理表URLを対象）
  python main.py cwv [--url URL]     # Core Web Vitals チェック
  python main.py content [--url URL] # 表記ゆれ・禁止表現チェック
  python main.py all                 # 全チェック実行
"""

import sys
import os
import time
import argparse
from pathlib import Path

# パス設定
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "checkers"))
sys.path.insert(0, str(BASE_DIR / "output"))

EXCEL_PATH = str(BASE_DIR.parent / "tjpコンテンツ管理表.xlsx")


def load_urls_from_excel(limit: int = 0) -> list[str]:
    """Excel管理表からURLを読み込む"""
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    ws = wb["運用サイトマップ"]
    urls = []
    for row in ws.iter_rows(min_row=6, values_only=True):
        u = row[7]
        if u and isinstance(u, str) and u.startswith("http"):
            urls.append(u.strip())
    if limit:
        urls = urls[:limit]
    return list(dict.fromkeys(urls))  # 重複除去


# ─────────────────────────────────────────
# コマンド: titles
# ─────────────────────────────────────────
def cmd_titles(args):
    print("=" * 60)
    print("▶ ページtitle/description 一括取得 & Excel更新")
    print("=" * 60)
    from checkers.content import fetch_titles_bulk
    from output.reporter import update_excel_titles

    urls = load_urls_from_excel()
    print(f"対象URL数: {len(urls)}")
    print("取得中（5並列・0.3s間隔）...\n")

    titles = fetch_titles_bulk(urls)
    fetched = sum(1 for v in titles.values() if v)
    print(f"\n取得完了: {fetched}/{len(urls)} 件")

    print("Excelに書き込み中...")
    updated = update_excel_titles(EXCEL_PATH, titles)
    print(f"✅ {updated} 行のページ説明を更新しました")


# ─────────────────────────────────────────
# コマンド: links
# ─────────────────────────────────────────
def cmd_links(args):
    print("=" * 60)
    print("▶ リンク切れチェック")
    print("=" * 60)
    from checkers.links import check_links
    from output.reporter import save_links_report

    if args.url:
        urls = [args.url]
    else:
        urls = load_urls_from_excel(args.limit or 0)

    print(f"対象URL数: {len(urls)}\n")
    results = check_links(urls, deep=args.deep)
    path = save_links_report(results)

    broken = [r for r in results if r.get("broken")]
    print(f"\n{'='*40}")
    print(f"✅ 正常: {len(results)-len(broken)} / ❌ リンク切れ: {len(broken)}")
    if broken:
        print("\n■ リンク切れ一覧:")
        for r in broken:
            print(f"  [{r['status']}] {r['url']}")
    print(f"\nレポート: {path}")


# ─────────────────────────────────────────
# コマンド: cwv
# ─────────────────────────────────────────
def cmd_cwv(args):
    print("=" * 60)
    print("▶ Core Web Vitals チェック（PageSpeed Insights API）")
    print("=" * 60)
    from checkers.cwv import check_cwv_bulk
    from output.reporter import save_cwv_report
    from config import PSI_API_KEY

    if not PSI_API_KEY:
        print("⚠️  PSI_API_KEY が未設定です。config.py に設定するとRate制限が緩和されます\n")

    if args.url:
        urls = [args.url]
    else:
        urls = load_urls_from_excel(args.limit or 20)

    strategy = args.strategy or "mobile"
    print(f"対象URL数: {len(urls)} / strategy: {strategy}\n")

    results = check_cwv_bulk(urls, strategy)
    path = save_cwv_report(results)

    ng = [r for r in results if r.get("ng_items")]
    print(f"\n{'='*40}")
    print(f"✅ KPI達成: {len(results)-len(ng)} / ❌ KPI未達: {len(ng)}")
    if ng:
        print("\n■ KPI未達ページ:")
        for r in ng:
            items = ", ".join(r["ng_items"])
            print(f"  {r['url']}")
            print(f"    LCP:{r.get('LCP')}ms  CLS:{r.get('CLS')}  INP:{r.get('INP')}ms  → [{items}]")
    print(f"\nレポート: {path}")


# ─────────────────────────────────────────
# コマンド: content
# ─────────────────────────────────────────
def cmd_content(args):
    print("=" * 60)
    print("▶ コンテンツ検証（表記ゆれ・禁止表現）")
    print("=" * 60)
    from checkers.content import check_content
    from output.reporter import save_content_report

    if args.url:
        urls = [args.url]
    else:
        urls = load_urls_from_excel(args.limit or 0)

    print(f"対象URL数: {len(urls)}\n")
    results = check_content(urls)
    path = save_content_report(results)
    print(f"\nレポート: {path}")


# ─────────────────────────────────────────
# コマンド: all
# ─────────────────────────────────────────
def cmd_all(args):
    print("=" * 60)
    print("▶ 全チェック実行")
    print("=" * 60)
    cmd_titles(args)
    print()
    cmd_links(args)
    print()
    cmd_cwv(args)
    print()
    cmd_content(args)


# ─────────────────────────────────────────
# CLI エントリポイント
# ─────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="toyota-check — toyota.jp 専用サイト検証ツール"
    )
    sub = parser.add_subparsers(dest="command")

    for cmd in ["titles", "links", "cwv", "content", "all"]:
        p = sub.add_parser(cmd)
        p.add_argument("--url", help="単一URLを指定（省略時はExcel管理表全件）")
        p.add_argument("--limit", type=int, help="URL数上限（テスト用）")
        if cmd == "links":
            p.add_argument("--deep", action="store_true", help="ページ内リンクも検査")
        if cmd == "cwv":
            p.add_argument("--strategy", choices=["mobile", "desktop"], default="mobile")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        return

    cmds = {
        "titles": cmd_titles,
        "links": cmd_links,
        "cwv": cmd_cwv,
        "content": cmd_content,
        "all": cmd_all,
    }
    cmds[args.command](args)


if __name__ == "__main__":
    main()
