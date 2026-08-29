# Ollama × Claude Code 統合

Ollamaのローカルモデル（Kimi K3など）をClaude Codeから使用するための統合実装。

## セットアップ手順

### 1. Ollamaをインストール・セットアップ（ローカル環境）

```bash
# Ollamaをインストール
# https://ollama.ai から最新版をダウンロード

# Kimi K3をプル
ollama pull kimi:k3

# Ollamaを起動
ollama serve
```

**注意：** Ollamaはデフォルトで `http://localhost:11434` で起動します

### 2. 統合サーバーを起動（このプロジェクト）

```bash
# 依存関係をインストール
npm install

# サーバーを起動
npm start

# または開発モード（ファイル変更時に自動再起動）
npm run dev
```

サーバーは `http://localhost:3000` で起動します。

### 3. APIテスト

ブラウザで `http://localhost:3000/api-test.html` にアクセス

またはコマンドラインから：

```bash
# サーバー状態確認
curl http://localhost:3000/health

# Kimi K3に質問
curl -X POST http://localhost:3000/api/ask \
  -H "Content-Type: application/json" \
  -d '{"question":"こんにちは"}'
```

### 4. 環境変数でカスタマイズ

```bash
# Ollamaのエンドポイントを変更
OLLAMA_URL=http://remote-server:11434 npm start

# 別のモデルを使用
OLLAMA_MODEL=mistral npm start
```

## API エンドポイント

### `GET /health`
サーバーとOllamaの接続状態を確認

**レスポンス例：**
```json
{
  "status": "ok",
  "ollama_url": "http://localhost:11434",
  "model": "kimi:k3"
}
```

### `POST /api/ask`
シンプルなテキスト質問

**リクエスト：**
```json
{"question": "あなたは誰ですか？"}
```

**レスポンス：**
```json
{
  "question": "あなたは誰ですか？",
  "answer": "私はKimi K3というAIです...",
  "model": "kimi:k3"
}
```

### `POST /api/chat`
チャット形式（複数ターンの会話）

**リクエスト：**
```json
{
  "messages": [
    {"role": "user", "content": "こんにちは"},
    {"role": "assistant", "content": "こんにちは！..."}
  ],
  "stream": false
}
```

### `GET /api/models`
利用可能なOllamaモデル一覧を取得

## 構成

- `server.js` - Node.js統合サーバー
- `api-test.html` - Webテストインターフェース
- `index.html` - 職業性ストレス簡易調査票（元のアプリ）

## ブランチ

`claude/ollama-claude-code-integration-ov612b` で開発中
