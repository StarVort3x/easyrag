from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from datetime import datetime

# 聊天相关模型
class ChatMessage(BaseModel):
    role: str  # user/assistant
    content: str
    type: Optional[str] = "text"  # text/file
    fileName: Optional[str] = None

class ChatRequest(BaseModel):
    knowledgeId: int
    content: str
    chatId: Optional[str] = None  # 对话ID，用于多对话支持

class ChatResponse(BaseModel):
    code: int = 200
    msg: str = "success"
    data: Dict[str, Any]

class ChatDetailedResponse(BaseModel):
    """详细的聊天响应，包含思考过程和引用文献"""
    code: int = 200
    msg: str = "success"
    data: Dict[str, Any]  # {reply, thinking, references}

class ChatHistoryResponse(BaseModel):
    code: int = 200
    msg: str = "success"
    data: List[ChatMessage]

# 知识库相关模型
class KnowledgeBase(BaseModel):
    id: int
    label: str
    children: Optional[List['KnowledgeBase']] = None

class KnowledgeListResponse(BaseModel):
    code: int = 200
    msg: str = "success"
    data: List[KnowledgeBase]

class AddKnowledgeRequest(BaseModel):
    name: str

class AddKnowledgeResponse(BaseModel):
    code: int = 200
    msg: str = "success"
    data: Dict[str, Any]

# 文件相关模型
class FileUploadResponse(BaseModel):
    code: int = 200
    msg: str = "success"
    data: Dict[str, Any]

# 多对话相关模型
class CreateChatRequest(BaseModel):
    """创建对话请求"""
    title: str
    knowledgeId: int

class ChatInfo(BaseModel):
    """对话信息"""
    chatId: str
    title: str
    knowledgeId: int
    createdAt: str
    updatedAt: str

class CreateChatResponse(BaseModel):
    """创建对话响应"""
    code: int = 200
    msg: str = "success"
    data: ChatInfo

class DeleteChatRequest(BaseModel):
    """删除对话请求"""
    chatId: str

class CommonResponse(BaseModel):
    """通用响应"""
    code: int = 200
    msg: str = "success"
    data: Dict[str, Any] = {}

