# Excel AI

Local-first AI assistant for Excel financial models. Runs entirely on your machine — no data leaves your network.

## Prerequisites

- Docker Desktop (running)
- Node.js 18+
- 8GB RAM minimum
- NVIDIA GPU optional (speeds up Ollama)

## Quick Start

```bash
git clone <repo>
cd excel-ai
./scripts/setup.sh          # starts backend stack + pulls LLM (~5-10 min first run)
cd frontend && npm install && npm start
```

Open Excel → Insert → Add-ins → Upload My Add-in → select `frontend/manifest.xml`

## How to Use

- Drop an `.xlsx` file into the upload zone in the task pane
- Ask questions in natural language: "What drives NPV?", "Why is IRR 47%?", "What does D5 do?"
- Review and approve the execution plan before any writes are made
- Switch to the Anomalies tab to scan for statistical outliers

## Project Structure

```
excel-ai/
  backend/
    parser/       xlsx_parser.py  graph_builder.py
    rag/          chunker.py  local_embedder.py  chroma_store.py  retrieval.py
    agent/        ollama_client.py  excel_agent.py  tools.py
    api/          main.py  routes/  models.py  dependencies.py
    db/           connection.py  migrations/
    config/       settings.py
    tests/
  frontend/
    src/
      components/ TaskPane/  Chat/  Anomalies/  shared/
      hooks/      useSSE.ts  useWorkbook.ts  useExcelEvents.ts
      services/   api.ts  excel.ts
      types/      index.ts
  scripts/        setup.sh  reset.sh  test_api.sh
  docker-compose.yml
```

## Tech Stack

| Component | Technology | Why |
|---|---|---|
| LLM | Ollama llama3.1:8b | Fully local, no data egress |
| Embeddings | sentence-transformers paraphrase-multilingual-mpnet-base-v2 | Local, dim=768, multilingual |
| Vector DB | ChromaDB | In-process, persistent, cosine similarity |
| Backend | FastAPI + Python 3.11 | Async SSE streaming, dependency injection |
| Graph | networkx | PageRank, topological sort, Louvain clustering |
| Anomalies | scikit-learn IsolationForest | Unsupervised, no labelled data needed |
| Frontend | React 18 + TypeScript + Office.js | Only supported Excel Add-in API |
| Database | PostgreSQL 16 + asyncpg | Session persistence, JSONB for snapshots |

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `DATABASE_URL` | `postgresql+asyncpg://postgres:postgres@localhost:5432/excelai` | PostgreSQL connection |
| `OLLAMA_BASE_URL` | `http://localhost:11434` | Ollama API endpoint |
| `OLLAMA_MODEL` | `llama3.1:8b` | Primary LLM model |
| `OLLAMA_FALLBACK_MODEL` | `mistral` | Fallback if primary unavailable |
| `EMBEDDING_MODEL` | `paraphrase-multilingual-mpnet-base-v2` | Sentence-transformers model |
| `MODELS_DIR` | `./models` | Local model cache directory |
| `CHROMA_PATH` | `./chroma_db` | ChromaDB persistence directory |
| `LOG_LEVEL` | `INFO` | Logging level |

## Dev Without Docker

```bash
# Terminal 1 — PostgreSQL
docker run -p 5432:5432 -e POSTGRES_PASSWORD=postgres -e POSTGRES_DB=excelai postgres:16-alpine

# Terminal 2 — Ollama
ollama serve
ollama pull llama3.1:8b

# Terminal 3 — Backend
cd backend
pip install -r requirements.txt
cp ../.env.example ../.env
uvicorn api.main:app --reload --port 8000

# Terminal 4 — Frontend
cd frontend
npm install && npm start
```

## Troubleshooting

- **Ollama not running**: `GET /health` will show `ollama: offline`. Run `ollama serve` or start Docker container.
- **Model slow**: First inference after startup warms up the model. Subsequent calls are faster. GPU passthrough recommended.
- **Chroma error**: Delete `./chroma_db/` and re-upload. Run `./scripts/reset.sh` for full reset.
- **Office.js not loading**: Ensure `npm start` is running on `https://localhost:3000` (HTTPS required). Trust the self-signed certificate in your browser first.
