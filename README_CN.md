# EasyRAG - 本地 RAG 系统

一个基于 Ollama 的本地 RAG（检索增强生成）系统，包含前端界面和 FastAPI 后端。

## 🚀 快速开始

### 1. 安装依赖

```bash
# 后端
cd backend
pip install -r requirements.txt

# 前端
cd ../ruoyi-ui
npm install
```

### 2. 启动服务

**后端**（终端1）：
```bash
cd backend
python start.py
```

**前端**（终端2）：
```bash
cd ruoyi-ui
npm run serve
```

### 3. 打开浏览器

访问 http://localhost:8080

## ✨ 功能特性

- ✅ 基于 Ollama 的本地 LLM 支持
- ✅ 文档上传和识别（PDF、TXT、Markdown）
- ✅ 知识库管理
- ✅ 对话历史管理（SQLite 存储）
- ✅ 查询重写优化
- ✅ 向量检索和相似度匹配
- ✅ 思考过程显示
- ✅ 引用文献显示

## 📋 系统要求

- Python 3.8+
- Node.js 14+
- Ollama（已安装）

## 🔧 配置

编辑 `backend/config.yaml`：

```yaml
ollama_model: "qwen2:7b"           # 使用的模型
top_k: 7                           # 检索结果数量
ollama_api:
  base_url: "http://localhost:11434/v1"
  api_key: "qwen2:7b"
```

## 📁 项目结构

```
easyrag/
├── backend/                    # FastAPI 后端
│   ├── app/
│   │   ├── api/               # API 路由
│   │   ├── services/          # 业务逻辑
│   │   ├── models/            # 数据模型
│   │   └── database.py        # SQLite 数据库
│   ├── config.yaml            # 配置文件
│   └── requirements.txt       # Python 依赖
│
├── ruoyi-ui/                  # Vue 前端
│   ├── src/
│   │   ├── views/rag/        # RAG 页面
│   │   ├── components/       # 组件
│   │   └── api/              # API 调用
│   └── package.json          # NPM 依赖
│
└── README_CN.md              # 本文件
```

## 🔌 API 端点

### 聊天
- `POST /api/rag/chat` - 发送消息
- `GET /api/rag/chat/history?knowledgeId=1` - 获取对话历史
- `POST /api/rag/chat/clear?knowledgeId=1` - 清空对话

### 知识库
- `GET /api/rag/knowledge/list` - 获取知识库列表
- `POST /api/rag/knowledge/add` - 新增知识库

### 文件
- `POST /api/rag/file/upload` - 上传文件
- `GET /api/rag/file/list` - 获取文件列表

## 💾 数据存储

对话历史自动保存到 SQLite 数据库：
```
backend/chat_history.db
```

## 🐛 常见问题

**Q: 后端启动失败**
- 检查 Python 版本 (需要 3.8+)
- 运行 `pip install -r requirements.txt`

**Q: 前端无法连接后端**
- 确保后端运行在 http://0.0.0.0:8000
- 检查防火墙设置

**Q: 对话很慢**
- 这是正常的，模型推理需要时间
- 确保 Ollama 服务正在运行

## 📝 许可证

请查看各子项目的 LICENSE 文件。
