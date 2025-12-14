# Lark Voice Agent

音声でLarkのタスクを管理するAIエージェント

## 機能

- **音声入力**: macOSのDictation機能を使用した音声認識
- **タスク作成**: 自然言語でタスクを作成
- **サブタスク分解**: 複雑なタスクを自動的にサブタスクに分解
- **タスク完了**: 音声でタスクを完了としてマーク
- **タスク一覧**: 現在のタスクを表示

## セットアップ

### 1. 依存関係のインストール

```bash
cd ~/Projects/lark-voice-agent
pip install -e .
```

### 2. 環境変数の設定

```bash
cp .env.example .env
```

`.env`ファイルを編集して以下の値を設定:

```
LARK_APP_ID=your_app_id
LARK_APP_SECRET=your_app_secret
ANTHROPIC_API_KEY=your_anthropic_api_key
```

### 3. Lark APIの設定

1. [Lark Open Platform](https://open.larksuite.com/)にアクセス
2. アプリを作成
3. 以下の権限を付与:
   - `task:task:read` - タスクの読み取り
   - `task:task:write` - タスクの書き込み
   - `task:tasklist:read` - タスクリストの読み取り

### 4. macOS Dictationの設定（音声入力モード使用時）

1. システム設定 → キーボード → 音声入力
2. 音声入力をオンに
3. 言語: 日本語を追加
4. ショートカット: Fnキーを2回押す（推奨）

## 使い方

### 起動

```bash
python -m src.main
```

または:

```bash
lark-agent
```

### 入力モード

1. **音声入力モード**: Fnキーを2回押して話しかける
2. **テキスト入力モード**: キーボードで入力

### コマンド例

```
「買い物リストを作成して」
→ 新しいタスク「買い物リスト」を作成

「ウェブサイトリニューアルのタスクを分解して」
→ 親タスクと具体的なサブタスクを自動生成

「レポート作成を完了」
→ 該当タスクを完了としてマーク

「タスク一覧を見せて」
→ 現在のタスクを表示
```

## アーキテクチャ

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│   音声入力   │────▶│  LLMエージェント │────▶│  Lark API   │
│ (Dictation) │     │   (Claude)    │     │   (Tasks)   │
└─────────────┘     └──────────────┘     └─────────────┘
                           │
                    タスク分析・分解
```

## ファイル構成

```
lark-voice-agent/
├── src/
│   ├── __init__.py
│   ├── main.py           # メインアプリケーション
│   ├── agent.py          # LLMエージェント
│   ├── lark_client.py    # Lark API クライアント
│   └── speech_recognizer.py  # 音声認識
├── .env.example          # 環境変数テンプレート
├── pyproject.toml        # プロジェクト設定
└── README.md
```

## トラブルシューティング

### 「Speech recognition not authorized」エラー

1. システム設定 → プライバシーとセキュリティ → 音声認識
2. ターミナル（または使用しているアプリ）を許可

### Lark API エラー

1. App IDとApp Secretが正しいか確認
2. アプリの権限設定を確認
3. アプリが公開されているか確認

## ライセンス

MIT
