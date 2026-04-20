from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api import rag

app = FastAPI(title="Easy RAG API", version="1.0.0")

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境应限制具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(rag.router, prefix="/api/rag", tags=["RAG"])

@app.get("/")
async def root():
    return {"message": "Easy RAG API", "version": "1.0.0"}

@app.get("/health")
async def health():
    return {"status": "healthy"}

