import torch
import ollama
import os
import json
import yaml
from openai import OpenAI
from typing import List, Dict, Optional
import re
import PyPDF2
from pathlib import Path
from app.database import chat_db

class RAGService:
    def __init__(self, config_path: str = None):
        """初始化 RAG 服务"""
        self.config = self.load_config(config_path)
        self.client = OpenAI(
            base_url=self.config["ollama_api"]["base_url"],
            api_key=self.config["ollama_api"]["api_key"]
        )
        self.vault_content = []
        self.vault_embeddings_tensor = None
        self.conversation_histories: Dict[int, List[Dict]] = {}  # 按知识库ID存储对话历史
        
        # 加载知识库
        self.load_vault()
    
    def load_config(self, config_path: str = None) -> dict:
        """加载配置文件"""
        if config_path is None:
            # 尝试多个可能的路径
            possible_paths = [
                "easy-local-rag-main/config.yaml",
                "../easy-local-rag-main/config.yaml",
                "config.yaml"
            ]
            for path in possible_paths:
                if os.path.exists(path):
                    config_path = path
                    break
        
        if config_path and os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as file:
                    return yaml.safe_load(file)
            except Exception as e:
                print(f"Error loading config file: {e}")
        
        # 使用默认配置
            # 使用默认配置
            return {
                "vault_file": "vault.txt",
                "embeddings_file": "vault_embeddings.json",
                "ollama_model": "qwen2:7b",
                "top_k": 7,
                "system_message": "You are a helpful assistant that is an expert at extracting the most useful information from a given text",
                "ollama_api": {
                    "base_url": "http://localhost:11434/v1",
                    "api_key": "qwen2:7b"
                }
            }
    
    def load_vault(self):
        """加载知识库内容并生成嵌入向量"""
        vault_file = self.config.get("vault_file", "vault.txt")
        embeddings_file = self.config.get("embeddings_file", "vault_embeddings.json")
        
        # 加载内容
        if os.path.exists(vault_file):
            with open(vault_file, "r", encoding='utf-8') as f:
                self.vault_content = [line.strip() for line in f.readlines() if line.strip()]
        
        # 加载或生成嵌入向量
        if os.path.exists(embeddings_file) and len(self.vault_content) > 0:
            try:
                with open(embeddings_file, "r", encoding="utf-8") as f:
                    embeddings = json.load(f)
                    if len(embeddings) == len(self.vault_content):
                        self.vault_embeddings_tensor = torch.tensor(embeddings)
                    else:
                        self.generate_embeddings(embeddings_file)
            except:
                self.generate_embeddings(embeddings_file)
        elif len(self.vault_content) > 0:
            self.generate_embeddings(embeddings_file)
        else:
            self.vault_embeddings_tensor = torch.tensor([])
    
    def generate_embeddings(self, embeddings_file: str):
        """生成嵌入向量"""
        embeddings = []
        for content in self.vault_content:
            try:
                response = ollama.embeddings(model='mxbai-embed-large', prompt=content)
                embeddings.append(response["embedding"])
            except Exception as e:
                print(f"Error generating embedding: {e}")
                embeddings.append([0.0] * 1024)  # 默认维度
        
        self.vault_embeddings_tensor = torch.tensor(embeddings)
        
        # 保存嵌入向量
        try:
            with open(embeddings_file, "w", encoding="utf-8") as f:
                json.dump(embeddings, f)
        except Exception as e:
            print(f"Error saving embeddings: {e}")
    
    def get_relevant_context(self, query: str, top_k: int = None) -> List[str]:
        """获取相关上下文"""
        if self.vault_embeddings_tensor is None or self.vault_embeddings_tensor.nelement() == 0:
            return []
        
        top_k = top_k or self.config.get("top_k", 7)
        
        try:
            # 生成查询的嵌入向量
            input_embedding = ollama.embeddings(model='mxbai-embed-large', prompt=query)["embedding"]
            input_tensor = torch.tensor(input_embedding).unsqueeze(0)
            
            # 检查维度是否匹配
            if input_tensor.shape[1] != self.vault_embeddings_tensor.shape[1]:
                print(f"Warning: Embedding dimension mismatch. Query: {input_tensor.shape[1]}, Vault: {self.vault_embeddings_tensor.shape[1]}")
                return []
            
            # 计算余弦相似度
            cos_scores = torch.cosine_similarity(
                input_tensor,
                self.vault_embeddings_tensor
            )
            
            # 获取 top-k
            top_k = min(top_k, len(cos_scores))
            if top_k <= 0:
                return []
            top_indices = torch.topk(cos_scores, k=top_k)[1].tolist()
            
            # 返回相关上下文
            return [self.vault_content[idx].strip() for idx in top_indices]
        except Exception as e:
            print(f"Error getting relevant context: {e}")
            return []
    
    def rewrite_query(self, user_input: str, conversation_history: List[Dict]) -> str:
        """重写查询以改善检索"""
        if len(conversation_history) <= 1:
            return user_input
        
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
        
        try:
            response = self.client.chat.completions.create(
                model=self.config["ollama_model"],
                messages=[{"role": "system", "content": prompt}],
                max_tokens=200,
                n=1,
                temperature=0.1,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"Error rewriting query: {e}")
            return user_input
    
    def chat(self, knowledge_id: int, user_input: str, use_rewrite: bool = True, chat_id: Optional[str] = None) -> dict:
        """处理聊天请求，返回详细信息包括思考过程和引用文献"""
        # 获取或创建对话历史
        if knowledge_id not in self.conversation_histories:
            self.conversation_histories[knowledge_id] = []
        
        conversation_history = self.conversation_histories[knowledge_id]
        
        # 重写查询（如果需要）
        if use_rewrite and len(conversation_history) > 0:
            rewritten_query = self.rewrite_query(user_input, conversation_history)
        else:
            rewritten_query = user_input
        
        # 获取相关上下文
        relevant_context = self.get_relevant_context(rewritten_query)
        
        # 构建带上下文的用户输入
        user_input_with_context = user_input
        context_str = ""
        if relevant_context:
            context_str = "\n".join(relevant_context)
            user_input_with_context = user_input + "\n\nRelevant Context:\n" + context_str
        
        # 添加到对话历史
        conversation_history.append({"role": "user", "content": user_input_with_context})
        
        # 构建消息
        messages = [
            {"role": "system", "content": self.config["system_message"]},
            *conversation_history
        ]
        
        # 调用模型
        try:
            response = self.client.chat.completions.create(
                model=self.config["ollama_model"],
                messages=messages,
                max_tokens=2000,
            )
            
            assistant_response = response.choices[0].message.content
            conversation_history.append({"role": "assistant", "content": assistant_response})
            
            # 构建详细响应
            # 生成思考过程描述
            thinking_process = self._generate_thinking_process(
                user_input, 
                rewritten_query, 
                use_rewrite, 
                len(conversation_history),
                len(relevant_context)
            )
            
            # 保存到数据库
            try:
                if chat_id:
                    # 使用新的多对话方法
                    chat_db.save_message_to_chat(chat_id, knowledge_id, "user", user_input)
                    chat_db.save_message_to_chat(
                        chat_id,
                        knowledge_id,
                        "assistant",
                        assistant_response,
                        thinking=thinking_process,
                        references=relevant_context
                    )
                else:
                    # 使用旧的方法（兼容性）
                    chat_db.create_session(knowledge_id)
                    chat_db.save_message(knowledge_id, "user", user_input)
                    chat_db.save_message(
                        knowledge_id, 
                        "assistant", 
                        assistant_response,
                        thinking=thinking_process,
                        references=relevant_context
                    )
            except Exception as e:
                print(f"Warning: Failed to save to database: {e}")
            
            return {
                "reply": assistant_response,
                "thinking": thinking_process,
                "references": relevant_context if relevant_context else []
            }
        except Exception as e:
            print(f"Error in chat: {e}")
            error_msg = f"抱歉，处理请求时出错：{str(e)}"
            return {
                "reply": error_msg,
                "thinking": "处理出错",
                "references": []
            }
    
    def _generate_thinking_process(self, user_input: str, rewritten_query: str, use_rewrite: bool, history_len: int, ref_count: int) -> str:
        """生成思考过程描述"""
        steps = []
        
        # 第一步：接收查询
        steps.append(f"1️⃣ 接收查询: \"{user_input}\"")
        
        # 第二步：查询处理
        if use_rewrite and history_len > 1:
            if user_input != rewritten_query:
                steps.append(f"2️⃣ 查询重写: \"{rewritten_query}\"")
            else:
                steps.append(f"2️⃣ 保持原始查询")
        else:
            steps.append(f"2️⃣ 直接使用原始查询")
        
        # 第三步：检索相关内容
        if ref_count > 0:
            steps.append(f"3️⃣ 检索相关文献: 找到 {ref_count} 条相关内容")
        else:
            steps.append(f"3️⃣ 未找到相关文献，使用通用知识回答")
        
        # 第四步：生成回复
        steps.append(f"4️⃣ 基于上下文生成回复")
        
        return " → ".join(steps)
    
    def clear_history(self, knowledge_id: int):
        """清空对话历史（数据库和内存）"""
        # 清空数据库
        try:
            chat_db.clear_history(knowledge_id)
        except Exception as e:
            print(f"Warning: Failed to clear database history: {e}")
        
        # 清空内存
        if knowledge_id in self.conversation_histories:
            self.conversation_histories[knowledge_id] = []
    
    def get_history(self, knowledge_id: int) -> List[Dict]:
        """获取对话历史（从数据库和内存中合并）"""
        # 先从数据库获取
        db_history = chat_db.get_history(knowledge_id)
        
        # 如果数据库中有记录，返回数据库记录
        if db_history:
            return db_history
        
        # 否则返回内存中的记录
        return self.conversation_histories.get(knowledge_id, [])
    
    def process_file(self, file_path: str, file_type: str) -> bool:
        """处理上传的文件并添加到知识库"""
        try:
            text = ""
            
            if file_type == "pdf":
                with open(file_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    for page in pdf_reader.pages:
                        if page.extract_text():
                            text += page.extract_text() + " "
            elif file_type in ["txt", "md"]:
                with open(file_path, 'r', encoding='utf-8') as f:
                    text = f.read()
            else:
                return False
            
            # 清理和分块
            text = re.sub(r'\s+', ' ', text).strip()
            sentences = re.split(r'(?<=[.!?]) +', text)
            chunks = []
            current_chunk = ""
            
            for sentence in sentences:
                if len(current_chunk) + len(sentence) + 1 < 1000:
                    current_chunk += (sentence + " ").strip()
                else:
                    chunks.append(current_chunk)
                    current_chunk = sentence + " "
            
            if current_chunk:
                chunks.append(current_chunk)
            
            # 追加到知识库
            vault_file = self.config.get("vault_file", "vault.txt")
            with open(vault_file, "a", encoding="utf-8") as vault_file_handle:
                for chunk in chunks:
                    vault_file_handle.write(chunk.strip() + "\n")
            
            # 重新加载知识库
            self.load_vault()
            
            return True
        except Exception as e:
            print(f"Error processing file: {e}")
            return False

# 全局 RAG 服务实例
rag_service = RAGService()

