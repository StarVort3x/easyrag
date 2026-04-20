<<<<<<< HEAD
# easyrag
=======
# Easy RAG 项目

一个基于 Ollama 的本地 RAG（检索增强生成）系统，包含前端界面和 FastAPI 后端。

## 项目结构

```
easyrag/
├── easy-local-rag-main/    # 原始 Python RAG 脚本
├── backend/                 # FastAPI 后端服务
├── ruoyi-ui/                # Vue.js 前端界面
├── nginx/                   # Nginx 配置
└── vault.txt               # 知识库文件
```

## 快速开始

### 1. 安装 Ollama

访问 https://ollama.com/download 下载并安装 Ollama。

### 2. 下载模型

```bash
ollama pull qwen3
ollama pull mxbai-embed-large
```

### 3. 启动后端服务

```bash
cd backend
pip install -r requirements.txt
python start.py
```

后端服务将在 http://localhost:8000 启动。

### 4. 启动前端服务

```bash
cd ruoyi-ui
npm install
npm run serve
```

前端服务将在 http://localhost:8080 启动。

### 5. 配置 Nginx（可选）

如果需要使用 Nginx 作为反向代理：

1. 修改 `nginx/conf/nginx.conf` 中的配置
2. 启动 Nginx：
   ```bash
   nginx/nginx.exe -c nginx/conf/nginx.conf
   ```

## 功能特性

- ✅ 基于 Ollama 的本地 LLM 支持
- ✅ 文档上传和识别（PDF、TXT、Markdown）
- ✅ 知识库管理
- ✅ 对话历史管理
- ✅ 查询重写优化
- ✅ 向量检索和相似度匹配

## API 文档

启动后端服务后，访问：
- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

## 开发说明

### 后端开发

后端使用 FastAPI 框架，主要文件：
- `backend/app/main.py` - 应用入口
- `backend/app/api/rag.py` - API 路由
- `backend/app/services/rag_service.py` - RAG 核心服务

### 前端开发

前端使用 Vue 2 + Element UI，主要文件：
- `ruoyi-ui/src/views/rag/chat/index.vue` - 聊天主页面
- `ruoyi-ui/src/api/` - API 接口定义

## 注意事项

1. 确保 Ollama 服务在运行
2. 首次使用需要生成嵌入向量，可能需要一些时间
3. 上传的文件会保存到 `uploads/` 目录
4. 知识库内容保存在 `vault.txt`

## 许可证

请查看各子项目的 LICENSE 文件。

>>>>>>> d601bacd (blue-first)
