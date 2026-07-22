# toyota-check

toyota.jp 専用のサイト検証・サイトマップ管理ツール（Streamlit製）。

## 機能

- 📄 タイトル取得 / 🔗 リンクチェック / 🌐 外部リンクチェック / 📌 リンク元調査
- ⚡ Core Web Vitals（PSI API） / 📝 表記ゆれ・禁止表現 / 🔍 アプリ・機能検出
- 🗂️ **サイトマップ管理** — コンテンツ管理表をマスタとして保持・編集・クロール差分検知・
  追加URL差分管理・陸の孤島チェック（詳細はアプリ内 📖 マニュアル）
- 🕷️ バックグラウンドクロール — マスタ／Excel／自動クロールを対象にジョブ実行

## 環境移行（git clone だけで再構築できます）

```bash
git clone https://github.com/mutulab/toyota-check.git
cd toyota-check
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml  # 値を記入
streamlit run app.py
```

- 動作要件: Python 3.10+（3.11推奨）
- サイトマップマスタは `data/sitemap_master.csv` に保持され、GitHubバックアップ有効時は
  リポジトリにも自動コミットされるため、clone しただけでマスタごと移行できます

## Secrets（.streamlit/secrets.toml）

| キー | 必須 | 説明 |
|---|---|---|
| `PASSWORD` | 推奨 | ログインパスワード（未設定時はコード内既定値） |
| `GITHUB_TOKEN` | 推奨 | マスタのGitHub自動バックアップ・復元用（Fine-grained PAT、対象リポジトリの Contents: Read and write） |
| `GITHUB_REPO` | 任意 | バックアップ先（既定: `mutulab/toyota-check`） |
| `PSI_API_KEY` | 任意 | Core Web Vitals 計測用（PageSpeed Insights API） |

secrets.toml は .gitignore 済み。トークンをコミットしないこと。

## デプロイ

- 本番: Streamlit Community Cloud が `main` ブランチを自動デプロイ（push後1〜2分）
- 本番URL・Secretsは Streamlit Cloud のアプリ設定（share.streamlit.io）で管理

## ドキュメント

- 利用者向け: アプリ内「📖 マニュアル」
- 引き継ぎ・保守者向け: [docs/HANDOVER.md](docs/HANDOVER.md)
