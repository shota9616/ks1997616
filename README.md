# AI経営管理ポータル

省力化投資補助金（一般型）の申請書類11種を自動生成するStreamlitアプリ。
マルチページ構成で、今後ツールを追加拡張できます。

## デプロイ

`main` ブランチにpushすると、Streamlit Cloudが自動でデプロイします。

## 編集ガイド

GitHub上で以下のファイルを編集すると、デプロイ後にアプリに反映されます。

| 変更したい内容 | 編集ファイル |
|--------------|------------|
| 事業計画の文章テンプレート | `scripts/content_generator.py` |
| 成長率・時給等の数値設定 | `scripts/config.py` |
| 業種別の工程テンプレート | `scripts/process_templates.py` |
| 図表の生成プロンプト | `scripts/diagram_generator.py` |
| AI臭除去のロジック | `scripts/auto_fix.py` |
| ポータルのホームページ | `app.py` |
| 省力化補助金ページのUI | `pages/1_省力化補助金申請.py` |
| 認証の仕組み | `lib/auth.py` |
| 共通スタイル | `lib/styles.py` |

## ファイル構成

```
app.py                             # ポータルホームページ（認証付き）
lib/
  auth.py                          # パスワード認証
  styles.py                        # 共通CSS・ヘッダー・フッター
pages/
  1_省力化補助金申請.py              # 省力化補助金 申請書類生成ツール
scripts/
  main.py                          # 再エクスポートハブ（後方互換性）
  models.py                        # データクラス定義
  config.py                        # 設定値（成長率、時給、業種テンプレ）
  hearing_reader.py                # ヒアリングシート読み込み
  process_templates.py             # 業種別Before/After工程テンプレート
  content_generator.py             # 事業計画テキスト生成（PREP法）
  diagram_generator.py             # 図表生成プロンプト（Gemini API）
  document_writer.py               # 事業計画書Part1-2 Word生成
  plan3_writer.py                  # 事業計画書Part3 Excel生成
  other_documents.py               # その他9種書類生成
  auto_fix.py                      # 自動修正ループ・AI臭除去
  transcription_to_hearing.py      # 議事録→ヒアリングシート変換
  validate.py                      # 品質スコアリング
  pdf_extractor.py                 # PDF読み取り（Gemini API）
templates/                         # 書類テンプレート（Excel/Word）
examples/                          # サンプルヒアリングシート
.streamlit/
  config.toml                      # テーマ・サーバー設定
```

## ローカル実行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Secrets設定

Streamlit Cloudの Secrets に以下を設定してください：

```toml
app_password = "ログインパスワード"
GEMINI_API_KEY = "your-gemini-api-key"
ANTHROPIC_API_KEY = "your-anthropic-api-key"
```
