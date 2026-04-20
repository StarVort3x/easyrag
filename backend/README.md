# Easy RAG Backend API

FastAPI 后端服务，提供 RAG（检索增强生成）功能。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动服务

### Windows
```bash
start.bat
```

### Linux/Mac
```bash
python start.py
```

或者使用 uvicorn 直接启动：
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

## API 文档

启动服务后，访问：
- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

## 配置说明

后端会自动查找配置文件 `easy-local-rag-main/config.yaml`，如果找不到会使用默认配置。

确保：
1. Ollama 服务已启动（默认端口 11434）
2. 已安装所需的模型：
   - `ollama pull qwen3` (或你配置的模型)
   - `ollama pull mxbai-embed-large`

## API 接口

### 聊天接口
- `POST /api/rag/chat` - 发送聊天消息
- `GET /api/rag/chat/history?knowledgeId={id}` - 获取对话历史
- `POST /api/rag/chat/clear?knowledgeId={id}` - 清空对话历史

### 知识库接口
- `GET /api/rag/knowledge/list` - 获取知识库列表
- `POST /api/rag/knowledge/add` - 新增知识库

### 文件接口
- `POST /api/rag/file/upload` - 上传并识别文件
- `GET /api/rag/file/list` - 获取已识别文件列表

