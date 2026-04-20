#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据库模块 - 管理对话历史的持久化存储
"""

import sqlite3
import json
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path

class ChatHistoryDB:
    """对话历史数据库管理"""
    
    def __init__(self, db_path: str = "chat_history.db"):
        """初始化数据库"""
        self.db_path = db_path
        self.init_db()
    
    def init_db(self):
        """初始化数据库表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建对话表 - 存储对话元数据
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT UNIQUE NOT NULL,
                title TEXT NOT NULL,
                knowledge_id INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建对话历史表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT NOT NULL,
                knowledge_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                thinking TEXT,
                ref_data TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (chat_id) REFERENCES chats(chat_id)
            )
        ''')
        
        # 创建对话会话表（兼容旧版本）
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chat_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                knowledge_id INTEGER NOT NULL UNIQUE,
                session_name TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建知识库表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS knowledge_bases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kb_id INTEGER NOT NULL UNIQUE,
                label TEXT NOT NULL,
                parent_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (parent_id) REFERENCES knowledge_bases(kb_id)
            )
        ''')
        
        # 创建索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_chat_id ON chat_history(chat_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_knowledge_id ON chat_history(knowledge_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_created_at ON chat_history(created_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_kb_id ON knowledge_bases(kb_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_parent_id ON knowledge_bases(parent_id)')
        
        conn.commit()
        conn.close()
    
    def save_message(self, knowledge_id: int, role: str, content: str, 
                    thinking: str = None, references: List[str] = None) -> int:
        """保存单条消息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        references_json = json.dumps(references) if references else None
        
        cursor.execute('''
            INSERT INTO chat_history (knowledge_id, role, content, thinking, ref_data)
            VALUES (?, ?, ?, ?, ?)
        ''', (knowledge_id, role, content, thinking, references_json))
        
        conn.commit()
        message_id = cursor.lastrowid
        conn.close()
        
        return message_id
    
    def get_history(self, knowledge_id: int, limit: int = 100) -> List[Dict]:
        """获取对话历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, role, content, thinking, ref_data, created_at
            FROM chat_history
            WHERE knowledge_id = ?
            ORDER BY created_at ASC
            LIMIT ?
        ''', (knowledge_id, limit))
        
        rows = cursor.fetchall()
        conn.close()
        
        messages = []
        for row in rows:
            msg = {
                "id": row[0],
                "role": row[1],
                "content": row[2],
                "thinking": row[3],
                "references": json.loads(row[4]) if row[4] else [],
                "created_at": row[5]
            }
            messages.append(msg)
        
        return messages
    
    def get_recent_history(self, knowledge_id: int, limit: int = 20) -> List[Dict]:
        """获取最近的对话历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, role, content, thinking, ref_data, created_at
            FROM chat_history
            WHERE knowledge_id = ?
            ORDER BY created_at DESC
            LIMIT ?
        ''', (knowledge_id, limit))
        
        rows = cursor.fetchall()
        conn.close()
        
        messages = []
        for row in reversed(rows):  # 反转以保持时间顺序
            msg = {
                "id": row[0],
                "role": row[1],
                "content": row[2],
                "thinking": row[3],
                "references": json.loads(row[4]) if row[4] else [],
                "created_at": row[5]
            }
            messages.append(msg)
        
        return messages
    
    def clear_history(self, knowledge_id: int) -> int:
        """清空对话历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM chat_history WHERE knowledge_id = ?', (knowledge_id,))
        
        conn.commit()
        deleted_count = cursor.rowcount
        conn.close()
        
        return deleted_count
    
    def delete_message(self, message_id: int) -> bool:
        """删除单条消息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM chat_history WHERE id = ?', (message_id,))
        
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        
        return success
    
    def get_history_count(self, knowledge_id: int) -> int:
        """获取对话历史数量"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT COUNT(*) FROM chat_history WHERE knowledge_id = ?', (knowledge_id,))
        count = cursor.fetchone()[0]
        conn.close()
        
        return count
    
    def export_history(self, knowledge_id: int, format: str = "json") -> str:
        """导出对话历史"""
        history = self.get_history(knowledge_id)
        
        if format == "json":
            return json.dumps(history, ensure_ascii=False, indent=2)
        elif format == "txt":
            lines = []
            for msg in history:
                role = "用户" if msg["role"] == "user" else "AI"
                lines.append(f"\n[{msg['created_at']}] {role}:")
                lines.append(msg["content"])
                if msg.get("thinking"):
                    lines.append(f"\n思考过程: {msg['thinking']}")
                if msg.get("references"):
                    lines.append(f"\n引用文献: {len(msg['references'])} 条")
            return "\n".join(lines)
        else:
            raise ValueError(f"Unsupported format: {format}")
    
    def create_session(self, knowledge_id: int, session_name: str = None) -> int:
        """创建对话会话"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        if session_name is None:
            session_name = f"Session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        try:
            cursor.execute('''
                INSERT INTO chat_sessions (knowledge_id, session_name)
                VALUES (?, ?)
            ''', (knowledge_id, session_name))
            conn.commit()
            session_id = cursor.lastrowid
        except sqlite3.IntegrityError:
            # 如果会话已存在，更新时间戳
            cursor.execute('''
                UPDATE chat_sessions
                SET updated_at = CURRENT_TIMESTAMP
                WHERE knowledge_id = ?
            ''', (knowledge_id,))
            conn.commit()
            cursor.execute('SELECT id FROM chat_sessions WHERE knowledge_id = ?', (knowledge_id,))
            session_id = cursor.fetchone()[0]
        
        conn.close()
        return session_id
    
    def get_session(self, knowledge_id: int) -> Optional[Dict]:
        """获取对话会话"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, knowledge_id, session_name, created_at, updated_at
            FROM chat_sessions
            WHERE knowledge_id = ?
        ''', (knowledge_id,))
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {
                "id": row[0],
                "knowledge_id": row[1],
                "session_name": row[2],
                "created_at": row[3],
                "updated_at": row[4]
            }
        return None
    
    # ==================== 多对话功能 ====================
    
    def create_chat(self, chat_id: str, title: str, knowledge_id: int) -> bool:
        """创建新对话"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO chats (chat_id, title, knowledge_id)
                VALUES (?, ?, ?)
            ''', (chat_id, title, knowledge_id))
            conn.commit()
            conn.close()
            return True
        except sqlite3.IntegrityError:
            conn.close()
            return False
    
    def get_chat(self, chat_id: str) -> Optional[Dict]:
        """获取对话信息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, chat_id, title, knowledge_id, created_at, updated_at
            FROM chats
            WHERE chat_id = ?
        ''', (chat_id,))
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {
                "id": row[0],
                "chat_id": row[1],
                "title": row[2],
                "knowledge_id": row[3],
                "created_at": row[4],
                "updated_at": row[5]
            }
        return None
    
    def delete_chat(self, chat_id: str) -> bool:
        """删除对话及其所有消息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # 删除对话的所有消息
            cursor.execute('DELETE FROM chat_history WHERE chat_id = ?', (chat_id,))
            # 删除对话本身
            cursor.execute('DELETE FROM chats WHERE chat_id = ?', (chat_id,))
            conn.commit()
            success = cursor.rowcount > 0
            conn.close()
            return success
        except Exception as e:
            conn.close()
            print(f"Error deleting chat: {e}")
            return False
    
    def save_message_to_chat(self, chat_id: str, knowledge_id: int, role: str, content: str,
                            thinking: str = None, references: List[str] = None) -> int:
        """保存消息到指定对话"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        references_json = json.dumps(references) if references else None
        
        try:
            cursor.execute('''
                INSERT INTO chat_history (chat_id, knowledge_id, role, content, thinking, ref_data)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (chat_id, knowledge_id, role, content, thinking, references_json))
            
            conn.commit()
            message_id = cursor.lastrowid
            conn.close()
            return message_id
        except Exception as e:
            conn.close()
            print(f"Error saving message: {e}")
            return -1
    
    def get_chat_history(self, chat_id: str, limit: int = 100) -> List[Dict]:
        """获取指定对话的历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, role, content, thinking, ref_data, created_at
            FROM chat_history
            WHERE chat_id = ?
            ORDER BY created_at ASC
            LIMIT ?
        ''', (chat_id, limit))
        
        rows = cursor.fetchall()
        conn.close()
        
        messages = []
        for row in rows:
            msg = {
                "id": row[0],
                "role": row[1],
                "content": row[2],
                "thinking": row[3],
                "references": json.loads(row[4]) if row[4] else [],
                "created_at": row[5]
            }
            messages.append(msg)
        
        return messages
    
    def clear_chat_history(self, chat_id: str) -> int:
        """清空指定对话的历史"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM chat_history WHERE chat_id = ?', (chat_id,))
        
        conn.commit()
        deleted_count = cursor.rowcount
        conn.close()
        
        return deleted_count
    
    def get_all_chats(self) -> List[Dict]:
        """获取所有对话列表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, chat_id, title, knowledge_id, created_at, updated_at
            FROM chats
            ORDER BY updated_at DESC
        ''')
        
        rows = cursor.fetchall()
        conn.close()
        
        chats = []
        for row in rows:
            chat = {
                "id": row[0],
                "chat_id": row[1],
                "title": row[2],
                "knowledge_id": row[3],
                "created_at": row[4],
                "updated_at": row[5]
            }
            chats.append(chat)
        
        return chats
    
    # ==================== 知识库管理 ====================
    
    def get_all_knowledge_bases(self) -> List[Dict]:
        """获取所有知识库（包括子知识库）"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 获取所有知识库
        cursor.execute('''
            SELECT kb_id, label, parent_id
            FROM knowledge_bases
            ORDER BY parent_id, kb_id
        ''')
        
        rows = cursor.fetchall()
        conn.close()
        
        # 构建嵌套结构
        kbs_dict = {}
        root_kbs = []
        
        for row in rows:
            kb_id, label, parent_id = row
            kb = {"id": kb_id, "label": label}
            kbs_dict[kb_id] = kb
            
            if parent_id is None:
                root_kbs.append(kb)
            else:
                if parent_id not in kbs_dict:
                    kbs_dict[parent_id] = {"id": parent_id, "label": "", "children": []}
                if "children" not in kbs_dict[parent_id]:
                    kbs_dict[parent_id]["children"] = []
                kbs_dict[parent_id]["children"].append(kb)
        
        return root_kbs
    
    def add_knowledge_base(self, kb_id: int, label: str, parent_id: Optional[int] = None) -> bool:
        """添加知识库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO knowledge_bases (kb_id, label, parent_id)
                VALUES (?, ?, ?)
            ''', (kb_id, label, parent_id))
            conn.commit()
            conn.close()
            return True
        except sqlite3.IntegrityError:
            conn.close()
            return False
    
    def delete_knowledge_base(self, kb_id: int) -> bool:
        """删除知识库及其子知识库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # 删除所有子知识库
            cursor.execute('DELETE FROM knowledge_bases WHERE parent_id = ?', (kb_id,))
            # 删除知识库本身
            cursor.execute('DELETE FROM knowledge_bases WHERE kb_id = ?', (kb_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            conn.close()
            return False


# 全局数据库实例
chat_db = ChatHistoryDB()
