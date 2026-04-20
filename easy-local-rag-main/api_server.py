import os
import json
import sqlite3
from datetime import datetime
from typing import Optional, Dict, List, Union

import torch
import ollama
from fastapi import FastAPI, UploadFile, File, Form, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openai import OpenAI
import PyPDF2
import re

PINK = '\033[95m'
CYAN = '\033[96m'
YELLOW = '\033[93m'
NEON_GREEN = '\033[92m'
RESET_COLOR = '\033[0m'


DB_PATH = "rag.db"


# ----------------- 复用 localrag 的核心逻辑 -----------------

def open_file(filepath: str) -> str:
    with open(filepath, 'r', encoding='utf-8') as infile:
        return infile.read()


def _append_chunks_to_vault(chunks: List[str], knowledge_id: int):
    """将文本块追加到对应知识库的 vault 文件"""
    vault_file = f"vault_{knowledge_id}.txt"
    with open(vault_file, "a", encoding="utf-8") as f:
        for chunk in chunks:
            # 过滤掉无法编码的字符
            clean_chunk = "".join(c for c in chunk if ord(c) < 0x10000)
            f.write(clean_chunk + "\n")


def _split_text_to_chunks(text: str, max_len: int = 1000) -> List[str]:
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []
    sentences = re.split(r"(?<=[.!?。！？]) +", text)
    chunks: List[str] = []
    current_chunk = ""
    for sentence in sentences:
        if len(current_chunk) + len(sentence) + 1 < max_len:
            current_chunk += (sentence + " ").strip()
        else:
            if current_chunk:
                chunks.append(current_chunk)
            current_chunk = sentence + " "
    if current_chunk:
        chunks.append(current_chunk)
    return chunks


def extract_chunks_from_pdf_bytes(data: bytes) -> List[str]:
    from io import BytesIO

    pdf_reader = PyPDF2.PdfReader(BytesIO(data))
    num_pages = len(pdf_reader.pages)
    text = ""
    for page_num in range(num_pages):
        page = pdf_reader.pages[page_num]
        page_text = page.extract_text() or ""
        text += page_text + " "
    chunks = _split_text_to_chunks(text)
    return chunks


def extract_chunks_from_text_bytes(data: bytes) -> List[str]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("utf-8", errors="ignore")
    chunks = _split_text_to_chunks(text)
    return chunks


def process_pdf_bytes(data: bytes, knowledge_id: int):
    """写入对应知识库的 vault 文件"""
    chunks = extract_chunks_from_pdf_bytes(data)
    _append_chunks_to_vault(chunks, knowledge_id)


def process_text_bytes(data: bytes, knowledge_id: int):
    """写入对应知识库的 vault 文件"""
    chunks = extract_chunks_from_text_bytes(data)
    _append_chunks_to_vault(chunks, knowledge_id)


def get_relevant_context(rewritten_input: str, vault_embeddings: torch.Tensor, vault_content: List[str], top_k: int = 3):
    if vault_embeddings.nelement() == 0:
        return []
    
    input_embedding = ollama.embeddings(model='mxbai-embed-large', prompt=rewritten_input)["embedding"]
    cos_scores = torch.cosine_similarity(torch.tensor(input_embedding).unsqueeze(0), vault_embeddings)
    
    # 获取相似度最高的索引和分数
    top_k = min(top_k, len(cos_scores))
    top_scores, top_indices = torch.topk(cos_scores, k=top_k)
    
    print(f"[DEBUG] Top {top_k} similarity scores: {top_scores.tolist()}")
    
    relevant_context = []
    seen = set()
    
    for i, idx in enumerate(top_indices.tolist()):
        content = vault_content[idx].strip()
        score = top_scores[i].item()
        
        # 只选择相似度 > 0.3 的内容
        if score < 0.3:
            print(f"[DEBUG] Skipping low similarity content: {score:.3f}")
            continue
            
        # 限制长度并去重
        truncated_content = content[:300] if len(content) > 300 else content
        if truncated_content not in seen:
            seen.add(truncated_content)
            relevant_context.append(truncated_content)
            print(f"[DEBUG] Added context chunk {i+1} (similarity: {score:.3f}): {truncated_content[:100]}...")
        
        if len(relevant_context) >= 2:  # 最多2个最相关的chunk
            break
    
    print(f"[DEBUG] Final context chunks: {len(relevant_context)}")
    return relevant_context


def rewrite_query(user_input_json: str, conversation_history: List[Dict], ollama_model: str, client: OpenAI) -> str:
    user_input = json.loads(user_input_json)["Query"]
    context = "\n".join([f"{msg['role']}: {msg['content']}" for msg in conversation_history[-2:]])
    prompt = f"""Rewrite the following query by incorporating relevant context from the conversation history.
    The rewritten query should:

    - Preserve the core intent and meaning of the original query
    - Expand and clarify the query to make it more specific and informative for retrieving relevant context
    - Avoid introducing new topics or queries that deviate from the original query
    - DONT EVER ANSWER the Original query, but instead focus on rephrasing and expanding it into a new query

    Return ONLY the rewritten query text, without any additional formatting or explanations.

    Conversation History:
    {context}

    Original query: [{user_input}]

    Rewritten query: 
    """
    response = client.chat.completions.create(
        model=ollama_model,
        messages=[{"role": "system", "content": prompt}],
        max_tokens=200,
        n=1,
        temperature=0.1,
    )
    rewritten_query = response.choices[0].message.content.strip()
    return json.dumps({"Rewritten Query": rewritten_query})


def ollama_chat(user_input: str,
                system_message: str,
                vault_embeddings: torch.Tensor,
                vault_content: List[str],
                ollama_model: str,
                conversation_history: List[Dict],
                client: OpenAI) -> str:
    conversation_history.append({"role": "user", "content": user_input})

    if len(conversation_history) > 1:
        query_json = {
            "Query": user_input,
            "Rewritten Query": ""
        }
        rewritten_query_json = rewrite_query(json.dumps(query_json), conversation_history, ollama_model, client)
        rewritten_query_data = json.loads(rewritten_query_json)
        rewritten_query = rewritten_query_data["Rewritten Query"]
        print(PINK + "Original Query: " + user_input + RESET_COLOR)
        print(PINK + "Rewritten Query: " + rewritten_query + RESET_COLOR)
    else:
        rewritten_query = user_input

    relevant_context = get_relevant_context(rewritten_query, vault_embeddings, vault_content)
    if relevant_context:
        context_str = "\n".join(relevant_context)
        print("Context Pulled from Documents: \n\n" + CYAN + context_str + RESET_COLOR)
    else:
        print(CYAN + "No relevant context found." + RESET_COLOR)

    user_input_with_context = user_input
    if relevant_context:
        user_input_with_context = user_input + "\n\nRelevant Context:\n" + context_str

    conversation_history[-1]["content"] = user_input_with_context

    messages = [
        {"role": "system", "content": system_message},
        *conversation_history
    ]

    print(f"[DEBUG] Sending messages to Ollama: {len(messages)} messages")
    print(f"[DEBUG] Last user message: {conversation_history[-1]['content'][:200]}...")

    try:
        response = client.chat.completions.create(
            model=ollama_model,
            messages=messages,
            max_tokens=2000,
        )
        answer = response.choices[0].message.content
        print(f"[DEBUG] Ollama response: {answer[:100]}...")
    except Exception as e:
        print(f"[DEBUG] Ollama API error: {e}")
        raise e

    conversation_history.append({"role": "assistant", "content": answer})

    return answer


# ----------------- FastAPI 应用与全局状态 -----------------

app = FastAPI()

# 允许前端跨域访问（根据需要可收紧域名）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ChatRequest(BaseModel):
    content: str
    knowledgeId: Optional[int] = None
    conversationId: Optional[str] = None


class KnowledgeCreate(BaseModel):
    name: str
    description: Optional[str] = None


class KnowledgeDelete(BaseModel):
    ids: List[int]


class FileDelete(BaseModel):
    fileIds: List[int]


# 会话历史：简单使用内存字典存储，key 为 conversationId，没有则使用 "default"
conversation_histories: Dict[str, List[Dict]] = {}


# ----------------- SQLite 工具函数 -----------------


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    # 知识库表
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS knowledge (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT,
            create_time TEXT
        )
        """
    )

    # 文件表
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS file (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            knowledge_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            size INTEGER,
            create_time TEXT,
            FOREIGN KEY (knowledge_id) REFERENCES knowledge(id) ON DELETE CASCADE
        )
        """
    )

    # 文本分片表
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS chunk (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            knowledge_id INTEGER NOT NULL,
            file_id INTEGER NOT NULL,
            content TEXT NOT NULL,
            embedding TEXT NOT NULL,
            FOREIGN KEY (knowledge_id) REFERENCES knowledge(id) ON DELETE CASCADE,
            FOREIGN KEY (file_id) REFERENCES file(id) ON DELETE CASCADE
        )
        """
    )

    # 默认知识库
    cur.execute("SELECT COUNT(1) AS cnt FROM knowledge")
    row = cur.fetchone()
    if row["cnt"] == 0:
        now = datetime.utcnow().isoformat()
        cur.execute(
            "INSERT INTO knowledge (name, description, create_time) VALUES (?, ?, ?)",
            ("默认知识库", "系统默认创建的知识库", now),
        )

    conn.commit()
    conn.close()


@app.on_event("startup")
async def startup_event():
    global client, vault_content, vault_embeddings_tensor, system_message, ollama_model
    global knowledge_vault_content, knowledge_vault_embeddings

    # 初始化数据库
    print(NEON_GREEN + "Initializing SQLite database..." + RESET_COLOR)
    init_db()

    print(NEON_GREEN + "Initializing Ollama API client..." + RESET_COLOR)
    client = OpenAI(
        base_url='http://localhost:11434/v1',
        api_key='qwen3'
    )

    # 模型名可从环境变量读取，默认 qwen2:7b
    ollama_model = os.getenv("OLLAMA_MODEL", "qwen3")

    # 1) 先从数据库加载各知识库的 chunk + embedding
    print(NEON_GREEN + "Loading chunks and embeddings from SQLite..." + RESET_COLOR)
    knowledge_vault_content = {}
    knowledge_vault_embeddings = {}

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT knowledge_id, content, embedding FROM chunk ORDER BY id")
    rows = cur.fetchall()
    conn.close()

    tmp_embeddings: Dict[int, List[List[float]]] = {}
    tmp_contents: Dict[int, List[str]] = {}
    for row in rows:
        k_id = row["knowledge_id"]
        content = row["content"]
        try:
            emb = json.loads(row["embedding"])
        except Exception:
            continue
        tmp_embeddings.setdefault(k_id, []).append(emb)
        tmp_contents.setdefault(k_id, []).append(content)

    for k_id, embs in tmp_embeddings.items():
        try:
            knowledge_vault_embeddings[k_id] = torch.tensor(embs)
            knowledge_vault_content[k_id] = tmp_contents.get(k_id, [])
        except Exception as e:
            print(f"Error converting embeddings for knowledge {k_id} to tensor: {e}")

    # 2) 加载各知识库的独立 vault 文件作为兜底
    print(NEON_GREEN + "Loading knowledge base vault files..." + RESET_COLOR)
    knowledge_vault_files = {}
    
    # 获取所有知识库ID
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id FROM knowledge")
    kb_rows = cur.fetchall()
    conn.close()
    
    for row in kb_rows:
        kb_id = row["id"]
        vault_file = f"vault_{kb_id}.txt"
        if os.path.exists(vault_file):
            with open(vault_file, "r", encoding='utf-8') as f:
                content = f.readlines()
                if content:  # 只有非空才加载向量
                    knowledge_vault_files[kb_id] = content
                    print(f"Loaded {len(content)} chunks from {vault_file}")
    
    # 3) 保持对全局 vault.txt 的兼容（作为最后的兜底）
    print(NEON_GREEN + "Loading global vault.txt (compatibility mode)..." + RESET_COLOR)
    vault_content = []
    if os.path.exists("vault.txt"):
        with open("vault.txt", "r", encoding='utf-8') as vault_file:
            vault_content = vault_file.readlines()

    if vault_content:
        print(NEON_GREEN + "Generating embeddings for global vault content..." + RESET_COLOR)
        vault_embeddings = []
        for content in vault_content:
            response = ollama.embeddings(model='mxbai-embed-large', prompt=content)
            vault_embeddings.append(response["embedding"])

        print("Converting embeddings to tensor...")
        vault_embeddings_tensor = torch.tensor(vault_embeddings) if vault_embeddings else torch.empty(0)
    else:
        vault_embeddings_tensor = torch.empty(0)

    system_message = "You are a helpful assistant that is an expert at extracting the most useful information from a given text. Also bring in extra relevant infromation to the user query from outside the given context."

    print(NEON_GREEN + "FastAPI RAG server started successfully" + RESET_COLOR)


@app.post("/api/rag/chat")
async def rag_chat(req: ChatRequest):
    if not req.content:
        return {"code": 400, "msg": "content 不能为空"}

    conv_id = req.conversationId or "default"
    history = conversation_histories.setdefault(conv_id, [])

    # 根据 knowledgeId 选择对应知识库的内容和向量
    kb_id = req.knowledgeId or 1
    kb_embeddings = knowledge_vault_embeddings.get(kb_id)
    kb_content = knowledge_vault_content.get(kb_id)

    # 如果数据库中没有，尝试从独立的 vault 文件加载
    if kb_embeddings is None or kb_embeddings.nelement() == 0:
        vault_file = f"vault_{kb_id}.txt"
        if os.path.exists(vault_file):
            with open(vault_file, "r", encoding='utf-8') as f:
                content = f.readlines()
                if content:
                    print(f"Loading embeddings from {vault_file} on-demand...")
                    embeddings = []
                    for line in content:
                        response = ollama.embeddings(model='mxbai-embed-large', prompt=line)
                        embeddings.append(response["embedding"])
                    kb_embeddings = torch.tensor(embeddings)
                    kb_content = content
                    # 缓存到内存
                    knowledge_vault_embeddings[kb_id] = kb_embeddings
                    knowledge_vault_content[kb_id] = kb_content

    # 最后回退到全局 vault.txt
    if kb_embeddings is None or kb_embeddings.nelement() == 0:
        kb_embeddings = vault_embeddings_tensor
        kb_content = vault_content

    try:
        answer = ollama_chat(
            user_input=req.content,
            system_message=system_message,
            vault_embeddings=kb_embeddings,
            vault_content=kb_content,
            ollama_model=ollama_model,
            conversation_history=history,
            client=client,
        )
    except Exception as e:
        print(f"Error in rag_chat: {e}")
        return {"code": 500, "msg": "服务器处理异常"}

    return {"code": 200, "data": {"reply": answer}}


@app.get("/api/rag/knowledge/list")
async def list_knowledge():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, name, description, create_time FROM knowledge ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    data = [dict(row) for row in rows]
    return {"code": 200, "data": data}


@app.post("/api/rag/knowledge/add")
async def add_knowledge(req: KnowledgeCreate):
    conn = get_db()
    cur = conn.cursor()
    now = datetime.utcnow().isoformat()
    cur.execute(
        "INSERT INTO knowledge (name, description, create_time) VALUES (?, ?, ?)",
        (req.name, req.description or "", now),
    )
    conn.commit()
    knowledge_id = cur.lastrowid
    conn.close()
    return {"code": 200, "data": {"id": knowledge_id, "name": req.name}}


@app.post("/api/rag/knowledge/delete")
async def delete_knowledge(req: KnowledgeDelete):
    if not req.ids:
        return {"code": 400, "msg": "ids 不能为空"}
    conn = get_db()
    cur = conn.cursor()
    placeholders = ",".join(["?"] * len(req.ids))
    cur.execute(f"DELETE FROM knowledge WHERE id IN ({placeholders})", req.ids)
    conn.commit()
    conn.close()
    return {"code": 200, "msg": "删除成功"}


@app.get("/api/rag/file/list")
async def list_files(knowledgeId: int = Query(..., alias="knowledgeId")):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT f.id, f.knowledge_id AS knowledgeId, f.filename, f.size, f.create_time
        FROM file f
        WHERE f.knowledge_id = ?
        ORDER BY f.id DESC
        """,
        (knowledgeId,),
    )
    rows = cur.fetchall()
    conn.close()
    data = [dict(row) for row in rows]
    return {"code": 200, "data": data}


@app.post("/api/rag/file/delete")
async def delete_files(req: FileDelete):
    if not req.fileIds:
        return {"code": 400, "msg": "fileIds 不能为空"}
    conn = get_db()
    cur = conn.cursor()
    placeholders = ",".join(["?"] * len(req.fileIds))
    # 先删除 chunk，再删 file
    cur.execute(f"DELETE FROM chunk WHERE file_id IN ({placeholders})", req.fileIds)
    cur.execute(f"DELETE FROM file WHERE id IN ({placeholders})", req.fileIds)
    conn.commit()
    conn.close()


@app.post("/api/rag/file/upload")
async def upload_files(files: Union[UploadFile, List[UploadFile]] = File(...), knowledgeId: Optional[str] = Form(None)):
    if not files:
        return {"code": 400, "msg": "files 不能为空"}

    # knowledgeId 允许为空，默认归入 1（默认知识库）
    try:
        kb_id = int(knowledgeId) if knowledgeId is not None else 1
    except ValueError:
        kb_id = 1

    # 先在内存中完成所有文件的抽取和分块，只有最终有可处理的文件时才写入 vault.txt 和数据库
    processed_files: List[str] = []
    skipped_files: List[str] = []
    pending_chunks: List[Tuple[str, List[str], bytes]] = []  # (filename, chunks, file_data)

    file_list = files if isinstance(files, list) else [files]
    for f in file_list:
        filename = f.filename or "unnamed"
        suffix = os.path.splitext(filename)[1].lower()
        data = await f.read()
        try:
            # 先抽取文本分片（不写库）
            if suffix == ".pdf":
                chunks = extract_chunks_from_pdf_bytes(data)
            elif suffix in [".txt", ".md"]:
                chunks = extract_chunks_from_text_bytes(data)
            else:
                continue

            if not chunks:
                continue

            pending_chunks.append((filename, chunks, data))
        except Exception as e:
            print(f"Error extracting text from file {filename}: {e}")
            continue

    if not pending_chunks:
        return {"code": 400, "msg": "未处理任何文件，当前仅支持 pdf/txt/md，或文件内容为空"}

    # 现在开始写入 vault.txt 和数据库
    conn = get_db()
    cur = conn.cursor()

    for filename, chunks, file_data in pending_chunks:
        # 去重：同一知识库中已存在同名文件，则本次不再写入 vault.txt 和数据库
        cur.execute(
            "SELECT id FROM file WHERE knowledge_id = ? AND filename = ? LIMIT 1",
            (kb_id, filename),
        )
        exists = cur.fetchone()
        if exists:
            skipped_files.append(filename)
            continue

        # 写入对应知识库的 vault 文件
        _append_chunks_to_vault(chunks, kb_id)

        # 写入数据库
        file_size = len(file_data)
        cur.execute(
            "INSERT INTO file (knowledge_id, filename, size, create_time) VALUES (?, ?, ?, ?)",
            (kb_id, filename, file_size, datetime.now().isoformat()),
        )
        file_id = cur.lastrowid
        for idx, chunk in enumerate(chunks):
            try:
                response = ollama.embeddings(model="mxbai-embed-large", prompt=chunk)
                emb = response.get("embedding", [])
                cur.execute(
                    "INSERT INTO chunk (file_id, knowledge_id, chunk_index, content, embedding) VALUES (?, ?, ?, ?, ?)",
                    (file_id, kb_id, idx, chunk, json.dumps(emb)),
                )
            except Exception as emb_err:
                print(f"Error generating embedding for chunk in file {filename}: {emb_err}")

        processed_files.append(filename)

    conn.commit()
    conn.close()

    # 文件写入 vault.txt 和数据库后，当前 chat 仍使用 vault.txt 的向量；
    # 如需让新内容参与检索，请重启服务或在后续版本中改为从数据库动态加载。
    return {
        "code": 200,
        "data": {
            "files": processed_files,
            "skippedFiles": skipped_files,
            "message": "文件已写入知识库（vault.txt + 数据库）。如需让新内容生效，请重新启动 RAG 服务以重新生成向量。同名文件不会重复写入。"
        }
    }


# 便于本地直接运行： python api_server.py
if __name__ == "__main__":
    import uvicorn

    uvicorn.run("api_server:app", host="0.0.0.0", port=8001, reload=True)
