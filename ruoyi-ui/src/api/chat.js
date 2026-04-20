import request from './request'

/**
 * 发送聊天消息
 * @param {Object} params - 请求参数
 * @param {string} params.content - 消息内容
 * @param {number} params.knowledgeId - 知识库ID
 * @param {string} [params.chatId] - 对话ID（可选）
 * @returns {Promise} 返回Promise对象
 */
export function sendChatMessage({ content, knowledgeId, chatId }) {
  if (!content || !knowledgeId) {
    return Promise.reject(new Error('消息内容和知识库ID不能为空'))
  }
  
  const data = {
    content,
    knowledgeId
  }
  
  if (chatId) {
    data.chatId = chatId
  }
  
  return request({
    url: '/api/rag/chat',
    method: 'post',
    data,
    showLoading: true
  })
}

/**
 * 获取对话历史
 * @param {number} knowledgeId - 知识库ID
 * @param {string} [conversationId] - 会话ID（可选）
 * @returns {Promise} 返回Promise对象
 */
export function getChatHistory(knowledgeId, conversationId) {
  if (!knowledgeId) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  const params = { knowledgeId }
  if (conversationId) {
    params.conversationId = conversationId
  }
  
  return request({
    url: '/api/rag/chat/history',
    method: 'get',
    params,
    showLoading: true
  })
}

/**
 * 清空对话历史
 * @param {number} knowledgeId - 知识库ID
 * @param {string} [conversationId] - 会话ID（可选）
 * @returns {Promise} 返回Promise对象
 */
export function clearChatHistory(knowledgeId, conversationId) {
  if (!knowledgeId) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  const data = { knowledgeId }
  if (conversationId) {
    data.conversationId = conversationId
  }
  
  return request({
    url: '/api/rag/chat/clear',
    method: 'post',
    data,
    showLoading: true
  })
}

/**
 * 创建新会话
 * @param {number} knowledgeId - 知识库ID
 * @param {string} title - 会话标题
 * @returns {Promise} 返回Promise对象
 */
export function createConversation(knowledgeId, title) {
  if (!knowledgeId || !title) {
    return Promise.reject(new Error('知识库ID和会话标题不能为空'))
  }
  
  return request({
    url: '/api/rag/conversation',
    method: 'post',
    data: {
      knowledgeId,
      title
    },
    showLoading: true
  })
}

/**
 * 获取会话列表
 * @param {number} knowledgeId - 知识库ID
 * @returns {Promise} 返回Promise对象
 */
export function getConversationList(knowledgeId) {
  if (!knowledgeId) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  return request({
    url: '/api/rag/conversation/list',
    method: 'get',
    params: { knowledgeId },
    showLoading: false
  })
}

// ==================== 多对话管理 API ====================

/**
 * 创建新对话
 * @param {Object} params - 请求参数
 * @param {string} params.title - 对话标题
 * @param {number} params.knowledgeId - 知识库ID
 * @returns {Promise} 返回Promise对象
 */
export function createChat({ title, knowledgeId }) {
  if (!title || !knowledgeId) {
    return Promise.reject(new Error('对话标题和知识库ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/create',
    method: 'post',
    data: {
      title,
      knowledgeId
    },
    showLoading: false
  })
}

/**
 * 删除对话
 * @param {string} chatId - 对话ID
 * @returns {Promise} 返回Promise对象
 */
export function deleteChat(chatId) {
  if (!chatId) {
    return Promise.reject(new Error('对话ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/delete',
    method: 'delete',
    params: { chatId },
    showLoading: false
  })
}

/**
 * 获取对话信息
 * @param {string} chatId - 对话ID
 * @returns {Promise} 返回Promise对象
 */
export function getChatInfo(chatId) {
  if (!chatId) {
    return Promise.reject(new Error('对话ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/info',
    method: 'get',
    params: { chatId },
    showLoading: false
  })
}

/**
 * 获取指定对话的历史
 * @param {string} chatId - 对话ID
 * @param {number} [limit] - 限制数量（默认100）
 * @returns {Promise} 返回Promise对象
 */
export function getChatHistoryByChat(chatId, limit = 100) {
  if (!chatId) {
    return Promise.reject(new Error('对话ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/history/by-chat',
    method: 'get',
    params: { chatId, limit },
    showLoading: false
  })
}

/**
 * 清空指定对话的历史
 * @param {string} chatId - 对话ID
 * @returns {Promise} 返回Promise对象
 */
export function clearChatHistoryByChat(chatId) {
  if (!chatId) {
    return Promise.reject(new Error('对话ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/clear-by-chat',
    method: 'post',
    params: { chatId },
    showLoading: false
  })
}

/**
 * 获取所有对话列表
 * @returns {Promise} 返回Promise对象
 */
export function getChatList() {
  return request({
    url: '/api/rag/chat/list',
    method: 'get',
    showLoading: false
  })
}

/**
 * 获取指定对话的消息历史
 * @param {string} chatId - 对话ID
 * @returns {Promise} 返回Promise对象
 */
export function getChatMessages(chatId) {
  if (!chatId) {
    return Promise.reject(new Error('对话ID不能为空'))
  }
  
  return request({
    url: '/api/rag/chat/messages',
    method: 'get',
    params: { chatId },
    showLoading: false
  })
}