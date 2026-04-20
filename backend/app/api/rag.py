from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Query, Body
from typing import List, Optional
from app.models.schemas import (
    ChatRequest, ChatResponse, ChatHistoryResponse,
    KnowledgeListResponse, AddKnowledgeRequest, AddKnowledgeResponse,
    FileUploadResponse, ChatMessage, CreateChatRequest, CreateChatResponse,
    ChatInfo, DeleteChatRequest, CommonResponse
)
from app.services.rag_service import rag_service
from app.database import chat_db
import os
import aiofiles
from pathlib import Path
from datetime import datetime

router = APIRouter()

# 聊天相关接口
@router.post("/chat", response_model=ChatResponse)
async def chat(request: ChatRequest):
    """发送聊天消息"""
    try:
        response_content = rag_service.chat(
            knowledge_id=request.knowledgeId,
            user_input=request.content,
            use_rewrite=True,
            chat_id=request.chatId  # 传递 chat_id
        )
        # response_content 现在是一个字典，包含 reply, thinking, references
        return ChatResponse(data=response_content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/history", response_model=ChatHistoryResponse)
async def get_chat_history(knowledgeId: int = Query(...)):
    """获取对话历史"""
    try:
        history = rag_service.get_history(knowledgeId)
        messages = []
        for msg in history:
            messages.append(ChatMessage(
                role=msg["role"],
                content=msg["content"]
            ))
        return ChatHistoryResponse(data=messages)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/chat/clear")
async def clear_chat_history(knowledgeId: Optional[int] = Query(None), body: Optional[dict] = Body(None)):
    """清空对话历史"""
    try:
        # 支持从 Query 或 Body 获取参数
        if knowledgeId is None and body:
            knowledgeId = body.get("knowledgeId")
        if knowledgeId is None:
            raise HTTPException(status_code=400, detail="knowledgeId is required")
        rag_service.clear_history(knowledgeId)
        return {"code": 200, "msg": "success", "data": {}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 知识库相关接口
@router.get("/knowledge/list", response_model=KnowledgeListResponse)
async def get_knowledge_list():
    """获取知识库列表"""
    try:
        knowledge_bases = chat_db.get_all_knowledge_bases()
        return KnowledgeListResponse(data=knowledge_bases)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/knowledge/add", response_model=AddKnowledgeResponse)
async def add_knowledge(request: AddKnowledgeRequest):
    """新增知识库"""
    try:
        # 获取最大 ID
        existing_kbs = chat_db.get_all_knowledge_bases()
        max_id = 0
        def get_max_id(kbs):
            nonlocal max_id
            for kb in kbs:
                if kb["id"] > max_id:
                    max_id = kb["id"]
                if "children" in kb:
                    get_max_id(kb["children"])
        get_max_id(existing_kbs)
        new_id = max_id + 1
        
        # 添加到数据库
        success = chat_db.add_knowledge_base(new_id, request.name)
        if not success:
            raise HTTPException(status_code=400, detail="Failed to add knowledge base")
        
        return AddKnowledgeResponse(data={"id": new_id, "name": request.name})
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/knowledge/delete")
async def delete_knowledge(data: dict = Body(...)):
    """删除知识库"""
    try:
        ids = data.get("ids", [])
        if not ids:
            raise HTTPException(status_code=400, detail="ids is required")
        
        deleted_count = 0
        for kb_id in ids:
            if chat_db.delete_knowledge_base(kb_id):
                deleted_count += 1
        
        return {"code": 200, "msg": "success", "data": {"deletedCount": deleted_count}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 文件相关接口
@router.post("/file/upload", response_model=FileUploadResponse)
async def upload_file(
    files: List[UploadFile] = File(...),
    knowledgeId: int = Form(...)
):
    """上传并识别文件"""
    try:
        # 创建上传目录
        upload_dir = Path("uploads")
        upload_dir.mkdir(exist_ok=True)
        
        uploaded_files = []
        
        for file in files:
            # 检查文件类型
            file_ext = Path(file.filename).suffix.lower()
            allowed_extensions = ['.pdf', '.txt', '.md', '.doc', '.docx', '.xls', '.xlsx']
            
            if file_ext not in allowed_extensions:
                continue
            
            # 保存文件
            file_path = upload_dir / file.filename
            async with aiofiles.open(file_path, 'wb') as f:
                content = await file.read()
                await f.write(content)
            
            # 处理文件
            file_type = file_ext[1:]  # 去掉点号
            if file_type == 'pdf':
                success = rag_service.process_file(str(file_path), "pdf")
            elif file_type in ['txt', 'md']:
                success = rag_service.process_file(str(file_path), file_type)
            else:
                # 对于 Word/Excel 等格式，暂时跳过或使用其他库处理
                success = False
            
            if success:
                uploaded_files.append({
                    "name": file.filename,
                    "size": len(content)
                })
        
        return FileUploadResponse(data={
            "files": uploaded_files,
            "message": f"成功上传 {len(uploaded_files)} 个文件"
        })
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/file/list")
async def get_file_list(knowledgeId: int = Query(...)):
    """获取已识别文件列表"""
    try:
        upload_dir = Path("uploads")
        files = []
        
        if upload_dir.exists():
            for file_path in upload_dir.iterdir():
                if file_path.is_file():
                    stat = file_path.stat()
                    files.append({
                        "id": hash(file_path.name) % (10 ** 8),  # 生成简单的ID
                        "filename": file_path.name,
                        "size": stat.st_size,
                        "create_time": datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
                        "knowledgeId": knowledgeId
                    })
        
        return {"code": 200, "msg": "success", "data": files}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/file/delete")
async def delete_file(fileIds: dict = Body(...)):
    """删除文件"""
    try:
        ids = fileIds.get("fileIds", [])
        if not ids:
            raise HTTPException(status_code=400, detail="fileIds is required")
        
        upload_dir = Path("uploads")
        deleted_count = 0
        
        if upload_dir.exists():
            for file_path in upload_dir.iterdir():
                if file_path.is_file():
                    file_id = hash(file_path.name) % (10 ** 8)
                    if file_id in ids:
                        try:
                            file_path.unlink()
                            deleted_count += 1
                        except Exception as e:
                            print(f"Error deleting file {file_path}: {e}")
        
        return {"code": 200, "msg": "success", "data": {"deletedCount": deleted_count}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 对话历史相关接口
@router.get("/chat/history/count")
async def get_history_count(knowledgeId: int = Query(...)):
    """获取对话历史数量"""
    try:
        count = chat_db.get_history_count(knowledgeId)
        return {"code": 200, "msg": "success", "data": {"count": count}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/history/recent")
async def get_recent_history(knowledgeId: int = Query(...), limit: int = Query(20)):
    """获取最近的对话历史"""
    try:
        history = chat_db.get_recent_history(knowledgeId, limit)
        messages = []
        for msg in history:
            messages.append(ChatMessage(
                role=msg["role"],
                content=msg["content"]
            ))
        return {"code": 200, "msg": "success", "data": messages}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/chat/history/export")
async def export_history(knowledgeId: int = Query(...), format: str = Query("json")):
    """导出对话历史"""
    try:
        if format not in ["json", "txt"]:
            raise ValueError("Format must be 'json' or 'txt'")
        
        exported = chat_db.export_history(knowledgeId, format)
        
        if format == "json":
            return {"code": 200, "msg": "success", "data": {"content": exported, "format": "json"}}
        else:
            return {"code": 200, "msg": "success", "data": {"content": exported, "format": "txt"}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/chat/history/message")
async def delete_message(messageId: int = Query(...)):
    """删除单条消息"""
    try:
        success = chat_db.delete_message(messageId)
        if success:
            return {"code": 200, "msg": "success", "data": {"deleted": True}}
        else:
            raise HTTPException(status_code=404, detail="Message not found")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ==================== 多对话管理接口 ====================

@router.post("/chat/create", response_model=CreateChatResponse)
async def create_chat(request: CreateChatRequest):
    """创建新对话"""
    try:
        chat_id = str(int(datetime.now().timestamp() * 1000))
        success = chat_db.create_chat(chat_id, request.title, request.knowledgeId)
        
        if not success:
            raise HTTPException(status_code=400, detail="Failed to create chat")
        
        chat_info = chat_db.get_chat(chat_id)
        if not chat_info:
            raise HTTPException(status_code=500, detail="Chat created but not found")
        
        return CreateChatResponse(data=ChatInfo(
            chatId=chat_info["chat_id"],
            title=chat_info["title"],
            knowledgeId=chat_info["knowledge_id"],
            createdAt=str(chat_info["created_at"]),
            updatedAt=str(chat_info["updated_at"])
        ))
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/chat/delete")
async def delete_chat(chatId: str = Query(...)):
    """删除对话"""
    try:
        success = chat_db.delete_chat(chatId)
        if success:
            return {"code": 200, "msg": "success", "data": {"deleted": True}}
        else:
            raise HTTPException(status_code=404, detail="Chat not found")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/info")
async def get_chat_info(chatId: str = Query(...)):
    """获取对话信息"""
    try:
        chat_info = chat_db.get_chat(chatId)
        if not chat_info:
            raise HTTPException(status_code=404, detail="Chat not found")
        
        return {
            "code": 200,
            "msg": "success",
            "data": {
                "chatId": chat_info["chat_id"],
                "title": chat_info["title"],
                "knowledgeId": chat_info["knowledge_id"],
                "createdAt": str(chat_info["created_at"]),
                "updatedAt": str(chat_info["updated_at"])
            }
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/history/by-chat")
async def get_chat_history_by_chat(chatId: str = Query(...), limit: int = Query(100)):
    """获取指定对话的历史"""
    try:
        history = chat_db.get_chat_history(chatId, limit)
        messages = []
        for msg in history:
            messages.append({
                "id": msg["id"],
                "role": msg["role"],
                "content": msg["content"],
                "thinking": msg.get("thinking", ""),
                "references": msg.get("references", []),
                "createdAt": msg["created_at"]
            })
        return {"code": 200, "msg": "success", "data": messages}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/chat/clear-by-chat")
async def clear_chat_history_by_chat(chatId: str = Query(...)):
    """清空指定对话的历史"""
    try:
        deleted_count = chat_db.clear_chat_history(chatId)
        return {"code": 200, "msg": "success", "data": {"deletedCount": deleted_count}}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/list")
async def get_chat_list():
    """获取所有对话列表"""
    try:
        chats = chat_db.get_all_chats()
        chat_list = []
        for chat in chats:
            chat_list.append({
                "id": chat["chat_id"],
                "title": chat["title"],
                "knowledgeId": chat["knowledge_id"],
                "createdAt": str(chat["created_at"]),
                "updatedAt": str(chat["updated_at"])
            })
        return {"code": 200, "msg": "success", "data": chat_list}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/chat/messages")
async def get_chat_messages(chatId: str = Query(...)):
    """获取指定对话的消息历史"""
    try:
        if not chatId:
            raise HTTPException(status_code=400, detail="chatId is required")
        
        messages = chat_db.get_chat_history(chatId)
        message_list = []
        for msg in messages:
            message_list.append({
                "role": msg["role"],
                "content": msg["content"],
                "thinking": msg.get("thinking", ""),
                "references": msg.get("references", [])
            })
        return {"code": 200, "msg": "success", "data": message_list}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

