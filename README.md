# Ollama × Claude Code 統合

Ollamaのローカルモデルを使用するための統合実装。Web UIでインストール済みモデルから選んで質問できる。

## セットアップ手順

### 1. Ollamaをインストール・モデルをプル（ローカル環境）

```bash
# Ollamaをインストール
# https://ollama.ai から最新版をダウンロード

# 使いたいモデルをプル（複数プルしてOK、UIで切り替えられる）
ollama pull mistral
ollama pull llama3
ollama pull gemma2

# Ollamaを起動
ollama serve
```

**注意：** Ollamaはデフォルトで `http://localhost:11434` で起動します。
**Kimi K3はOllamaの公式レジストリには存在しないため `ollama pull` できません**（クラウド専用モデルのため）。Ollamaで使う場合は上記のような公開モデルを選んでください。

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

ブラウザで `http://localhost:3000/api-test.html` にアクセスすると、インストール済みのOllamaモデルがドロップダウンに一覧表示されるので、選んで質問できます。

またはコマンドラインから：

```bash
# サーバー状態確認
curl http://localhost:3000/health

# インストール済みモデル一覧
curl http://localhost:3000/api/models

# モデルを指定して質問（model省略時はデフォルトモデルを使用）
curl -X POST http://localhost:3000/api/ask \
  -H "Content-Type: application/json" \
  -d '{"question":"こんにちは","model":"mistral"}'
```

### 4. 環境変数でカスタマイズ

```bash
# Ollamaのエンドポイントを変更
OLLAMA_URL=http://remote-server:11434 npm start

# デフォルトモデルを変更（リクエストでmodelを指定しなかった場合に使用）
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
  "default_model": "mistral"
}
```

### `POST /api/ask`
シンプルなテキスト質問。`model` を省略するとデフォルトモデルを使用。

**リクエスト：**
```json
{"question": "あなたは誰ですか？", "model": "mistral"}
```

**レスポンス：**
```json
{
  "question": "あなたは誰ですか？",
  "answer": "私はAIアシスタントです...",
  "model": "mistral"
}
```

### `POST /api/chat`
チャット形式（複数ターンの会話）。`model` を省略するとデフォルトモデルを使用。

**リクエスト：**
```json
{
  "model": "mistral",
  "messages": [
    {"role": "user", "content": "こんにちは"},
    {"role": "assistant", "content": "こんにちは！..."}
  ],
  "stream": false
}
```

### `GET /api/models`
インストール済みのOllamaモデル一覧を取得。Web UIのドロップダウンはここから動的に生成される。

## 構成

- `server.js` - Node.js統合サーバー
- `api-test.html` - Webテストインターフェース
- `index.html` - 職業性ストレス簡易調査票（元のアプリ）

## ブランチ

`claude/ollama-claude-code-integration-ov612b` で開発中
