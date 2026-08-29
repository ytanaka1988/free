import express from 'express';
import axios from 'axios';

const app = express();
const PORT = process.env.PORT || 3000;
const OLLAMA_URL = process.env.OLLAMA_URL || 'http://localhost:11434';
const MODEL = process.env.OLLAMA_MODEL || 'kimi:k3';

app.use(express.json());

// ヘルスチェック
app.get('/health', (req, res) => {
  res.json({ status: 'ok', ollama_url: OLLAMA_URL, model: MODEL });
});

// Ollamaへのチャットリクエストをプロキシ
app.post('/api/chat', async (req, res) => {
  try {
    const { messages, stream = false } = req.body;

    const response = await axios.post(
      `${OLLAMA_URL}/api/chat`,
      {
        model: MODEL,
        messages: messages,
        stream: stream
      },
      {
        timeout: 60000,
        responseType: stream ? 'stream' : 'json'
      }
    );

    if (stream) {
      res.setHeader('Content-Type', 'text/event-stream');
      response.data.pipe(res);
    } else {
      res.json(response.data);
    }
  } catch (error) {
    console.error('Error calling Ollama:', error.message);
    res.status(500).json({
      error: error.message,
      ollama_url: OLLAMA_URL
    });
  }
});

// 利用可能なモデル一覧を取得
app.get('/api/models', async (req, res) => {
  try {
    const response = await axios.get(`${OLLAMA_URL}/api/tags`);
    res.json(response.data);
  } catch (error) {
    console.error('Error fetching models:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// シンプルなテストエンドポイント
app.post('/api/ask', async (req, res) => {
  try {
    const { question } = req.body;

    if (!question) {
      return res.status(400).json({ error: 'question is required' });
    }

    const response = await axios.post(
      `${OLLAMA_URL}/api/generate`,
      {
        model: MODEL,
        prompt: question,
        stream: false
      }
    );

    res.json({
      question: question,
      answer: response.data.response,
      model: MODEL
    });
  } catch (error) {
    console.error('Error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(PORT, () => {
  console.log(`Ollama integration server running on http://localhost:${PORT}`);
  console.log(`Ollama endpoint: ${OLLAMA_URL}`);
  console.log(`Model: ${MODEL}`);
});
