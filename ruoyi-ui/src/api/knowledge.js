import request from './request'

/**
 * 知识库信息
 * @typedef {Object} KnowledgeBase
 * @property {number} id - 知识库ID
 * @property {string} name - 知识库名称
 * @property {string} [description] - 知识库描述
 * @property {string} createTime - 创建时间
 * @property {string} updateTime - 更新时间
 * @property {number} [fileCount] - 文件数量
 * @property {number} [status] - 状态：0-禁用，1-启用
 */

/**
 * 获取知识库列表
 * @param {Object} [params] - 查询参数
 * @param {string} [params.keyword] - 搜索关键词
 * @param {number} [params.status] - 状态：0-禁用，1-启用
 * @param {number} [params.pageNum=1] - 页码
 * @param {number} [params.pageSize=10] - 每页数量
 * @returns {Promise<{list: KnowledgeBase[], total: number}>} 返回知识库列表和总数
 */
export function getKnowledgeList(params = {}) {
  const { 
    keyword, 
    status, 
    pageNum = 1, 
    pageSize = 10 
  } = params
  
  return request({
    url: '/api/rag/knowledge/list',
    method: 'get',
    params: {
      keyword,
      status,
      pageNum,
      pageSize
    },
    showLoading: true
  })
}

/**
 * 获取知识库详情
 * @param {number} id - 知识库ID
 * @returns {Promise<KnowledgeBase>} 返回知识库详情
 */
export function getKnowledgeDetail(id) {
  if (!id) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  return request({
    url: '/api/rag/knowledge/detail',
    method: 'get',
    params: { id },
    showLoading: true
  })
}

/**
 * 新增知识库
 * @param {Object} data - 知识库数据
 * @param {string} data.name - 知识库名称
 * @param {string} [data.description] - 知识库描述
 * @param {number} [data.status=1] - 状态：0-禁用，1-启用
 * @returns {Promise<{id: number}>} 返回新增的知识库ID
 */
export function addKnowledge(data) {
  if (!data || !data.name) {
    return Promise.reject(new Error('知识库名称不能为空'))
  }
  
  const payload = {
    name: data.name,
    description: data.description || '',
    status: data.status !== undefined ? data.status : 1
  }
  
  return request({
    url: '/api/rag/knowledge/add',
    method: 'post',
    data: payload,
    showLoading: true
  })
}

/**
 * 更新知识库
 * @param {Object} data - 知识库数据
 * @param {number} data.id - 知识库ID
 * @param {string} [data.name] - 知识库名称
 * @param {string} [data.description] - 知识库描述
 * @param {number} [data.status] - 状态：0-禁用，1-启用
 * @returns {Promise} 返回Promise对象
 */
export function updateKnowledge(data) {
  if (!data || !data.id) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  const payload = { ...data }
  
  return request({
    url: '/api/rag/knowledge/update',
    method: 'post',
    data: payload,
    showLoading: true
  })
}

/**
 * 删除知识库
 * @param {number|number[]} id - 知识库ID或ID数组
 * @returns {Promise} 返回Promise对象
 */
export function deleteKnowledge(id) {
  if (!id) {
    return Promise.reject(new Error('请选择要删除的知识库'))
  }
  
  const ids = Array.isArray(id) ? id : [id]
  
  return request({
    url: '/api/rag/knowledge/delete',
    method: 'post',
    data: { ids },
    showLoading: true
  })
}

/**
 * 更新知识库状态
 * @param {number|number[]} id - 知识库ID或ID数组
 * @param {number} status - 状态：0-禁用，1-启用
 * @returns {Promise} 返回Promise对象
 */
export function updateKnowledgeStatus(id, status) {
  if (!id) {
    return Promise.reject(new Error('请选择要操作的知识库'))
  }
  
  if (status === undefined || status === null) {
    return Promise.reject(new Error('状态不能为空'))
  }
  
  const ids = Array.isArray(id) ? id : [id]
  
  return request({
    url: '/api/rag/knowledge/updateStatus',
    method: 'post',
    data: { 
      ids, 
      status: Number(status) 
    },
    showLoading: true
  })
}

/**
 * 获取知识库统计信息
 * @returns {Promise<{total: number, active: number, fileCount: number, totalSize: string}>} 返回统计信息
 */
export function getKnowledgeStats() {
  return request({
    url: '/api/rag/knowledge/stats',
    method: 'get',
    showLoading: false
  })
}