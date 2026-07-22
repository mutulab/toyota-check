# toyota-check 引き継ぎ仕様書

最終更新: 2026-07-22 ／ 作成: TMC運用PMO

## 1. 概要

toyota.jp の品質検証（リンク・CWV・表記・アプリ検出）と、コンテンツ管理表（サイトマップ）の
マスタ管理を行う社内ツール。Streamlit 製の単一アプリで、チーム共有パスワードで保護。

- 本番: Streamlit Community Cloud（`main` ブランチ自動デプロイ、push後1〜2分）
- リポジトリ: https://github.com/mutulab/toyota-check
- 本番URL・Secrets管理: share.streamlit.io のアプリ設定画面

## 2. ファイル構成と責務

| パス | 責務 |
|---|---|
| `app.py` | エントリポイント。認証・サイドバー（チェック種別の選択）・各チェックのUI/実行・マニュアル |
| `sitemap_manager.py` | 🗂️ サイトマップ管理モードの全実装（下記 §4） |
| `crawler.py` | バックグラウンドクロールのジョブ管理（開始/照会/永続化、`reports/` に結果保存） |
| `fetcher.py` | HTML取得（UA偽装・文字コード判定）、リンク/リソース/メタ/テキスト抽出 |
| `detector.py` | アプリ・機能検出ロジック |
| `checkers/` | links / cwv / content 各チェック |
| `config.py` | UA・タイムアウト・CWV閾値・禁止表現/表記ゆれ辞書 |
| `dict_loader.py` + `hyoki_dict.xlsx` | 表記ゆれ辞書の読み込み |
| `main.py` | CLI版（レガシー。通常は未使用） |
| `data/` | サイトマップマスタ等の永続データ（§5） |
| `.streamlit/` | テーマ・secrets（secrets.tomlはgit管理外、.exampleを参照） |

## 3. 認証・Secrets

- ログイン: `check_password()`（app.py冒頭）。`st.secrets["PASSWORD"]`、未設定時は既定値
- Secrets一覧は README.md 参照。トークン類は必ず Streamlit Cloud の Secrets か
  ローカルの `.streamlit/secrets.toml`（git管理外）に置く

## 4. サイトマップ管理（sitemap_manager.py）仕様

### データモデル
- マスタ = tjpコンテンツ管理表「運用サイトマップ」シートの全列（見出し行=5行目、データ=6行目以降）
- 取り込み時にURL空行を除外し相対パスへ `https://toyota.jp` を補完
- 派生列（保存されず表示時に毎回計算・マスタ更新時のみ再計算のキャッシュあり）:
  - `階層1〜6`（パス分解、index.html は除去）、`階層深さ`、`種別判定`
- 管理列（マスタに保存される）: `追加日` `追加元`（クロール検知の追加行に記録）

### 種別判定ルール（classify_url）
- アプリ入口パス（TID移行対象_サービスのまとまり整理v1.0 突合）: `/service/` `/member/`
  `/login` `/profile` `/grade` `/cmpn` `/socialfes` `/follow` `/ucar_search` `/webservice`
  `/mailalert_service` `/measurement` `/faq/inquiry` → アプリ
- `/recall` はトップのみアプリ（配下の年別・campaignは静的届出ページ）
- 拡張子（pdf/画像等）→ ファイル、クエリ付き非HTML → アプリ、それ以外 → HTML（静的）
- アプリ入口パスが増えたら `APP_PREFIXES` を更新すること

### 主要機能
1. 取り込み/差し替え（Excel → マスタ保存）
2. 3ビュー: ツリー（build_tree、配下ページ数は接頭辞カウントでO(n×階層)）／
   一覧編集（ページング、保存で即マスタ反映）／ディレクトリ集計
3. クロール差分検知（crawl_discover）: BFS巡回→未登録URL検知→選択追加
   （URL正規化: フラグメント・index.html・末尾スラッシュ・httpsを同一視）
4. 追加URL差分管理: 追加日別サマリー・取り消し削除・CSV
5. 陸の孤島チェック（run_orphan_check）: マスタ全ページの被リンク集計→孤島検出、
   実行履歴20回分を保存し前回比差分を表示。**JS生成リンクは検出不可（判定は要検証）**
6. CSV/Excelダウンロード（CSVはBOM付きUTF-8）

### 性能
- 3万URL規模で検証済み（派生計算0.3s・ツリー0.4s・保存0.5s・マスタCSV約2MB）

## 5. データ永続化とバックアップ

| ファイル | 内容 |
|---|---|
| `data/sitemap_master.csv` | サイトマップマスタ本体 |
| `data/sitemap_meta.json` | 最終更新・操作・更新履歴（直近50件） |
| `data/orphan_history.json` | 孤島チェック実行履歴（直近20回） |

- 保存フロー: ローカル書き込み → `GITHUB_TOKEN` があれば GitHub Contents API で
  `data/sitemap_master.csv` を自動コミット（メッセージ `data: サイトマップマスタ更新（…）`）
- 復元フロー: ローカルに無ければ GitHub から自動取得（再デプロイ対策）
- 注意: Streamlit Cloud のローカルFSは再起動で消える。GitHubバックアップが実質の正本
- 同時編集の排他制御は無い（小規模チーム前提。最後の保存が勝つ）

## 6. バックグラウンドクロール

- URL取得方法: 🗂️ サイトマップマスタ（既定）／🕷️ 自動クロール／📊 Excel
- ジョブは `crawler.start_job()` でバックグラウンドスレッド実行、`reports/` に永続化。
  ブラウザを閉じても継続。「🔍 ジョブ確認」で照会
- 注意: Streamlit Cloud ではコンテナ再起動でジョブ・reportsが消える

## 7. 環境移行手順（git化）

1. `git clone` → `pip install -r requirements.txt`（Python 3.10+）
2. `.streamlit/secrets.toml.example` をコピーして値を記入
3. `streamlit run app.py` で起動確認
4. 新しいホスティングに載せる場合: リポジトリを新環境に接続し `main` を指定、
   Secrets を移設するだけ。マスタは `data/sitemap_master.csv`（リポジトリ内）から自動復元
5. GitHubリポジトリを移す場合: Secrets の `GITHUB_REPO` を新リポジトリに変更

## 8. 既知の制約・注意点

- toyota.jp は bot 対策あり。`config.py` のUA・リクエスト間隔（0.3s）を変えない
- 孤島チェック・全件クロールは1ページ約0.5秒（859件≒7分、3万件は数時間→上限指定を推奨）
- GitHub PAT には期限がある。失効すると保存時に黄色警告（理由表示）→ 再発行して差し替え
- Cookiebot・Kendra等の運用ルールは本ツールの対象外（運用設計書を参照）

## 9. 定型運用（推奨サイクル）

- 月次: サイトマップ管理でクロール差分検知 → 追加 → 孤島チェック → 差分をTQP定例で報告
- 都度: コンテンツ公開時に一覧編集でマスタ更新
- 四半期: マスタとExcel正本の突合（Excelダウンロード→比較）
