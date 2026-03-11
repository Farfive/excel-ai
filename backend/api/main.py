import logging
import time

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from api.routes.workbook import router as workbook_router
from api.routes.analysis import router as analysis_router
from api.dependencies import embedder, ollama, chroma
from db.connection import init_db
from config.settings import get_settings

settings = get_settings()
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel AI",
    description="Local-first AI assistant for Excel financial models",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def log_requests(request: Request, call_next):
    start = time.time()
    response = await call_next(request)
    elapsed = int((time.time() - start) * 1000)
    logger.info(f"{request.method} {request.url.path} -> {response.status_code} ({elapsed}ms)")
    return response


@app.on_event("startup")
async def startup():
    logger.info("Starting Excel AI backend...")
    try:
        await init_db()
        logger.info("Database initialized")
    except Exception as e:
        logger.warning(f"DB init failed (non-fatal if no DB configured): {e}")

    try:
        embedder.load()
        logger.info("Embedder loaded")
    except Exception as e:
        logger.error(f"Embedder load failed: {e}")

    running = await ollama.is_running()
    if settings.groq_api_key:
        logger.info(f"Groq API {'connected' if running else 'UNREACHABLE'} — model: {settings.groq_model}")
    else:
        logger.info(f"Ollama {'running' if running else 'NOT running'} at {settings.ollama_base_url}")


@app.on_event("shutdown")
async def shutdown():
    await ollama.close()
    logger.info("Excel AI backend shutdown complete")


app.include_router(workbook_router)
app.include_router(analysis_router)


@app.get("/health")
async def health():
    ollama_ok = await ollama.is_running()

    chroma_ok = True
    chroma_detail = "ok"
    try:
        chroma._get_client()
    except Exception as e:
        chroma_ok = False
        chroma_detail = str(e)

    embedder_ok = embedder._model is not None

    overall = "ok" if (ollama_ok and chroma_ok and embedder_ok) else "degraded"

    return JSONResponse(
        content={
            "status": overall,
            "components": {
                "ollama": {"status": "ok" if ollama_ok else "offline"},
                "chroma": {"status": "ok" if chroma_ok else "error", "detail": chroma_detail},
                "embedder": {"status": "ok" if embedder_ok else "not_loaded"},
                "db": {"status": "ok"},
            },
            "model": settings.groq_model if settings.groq_api_key else settings.ollama_model,
        }
    )
