import request from './request'

/**
 * 文件上传
 * @param {Object} params - 上传参数
 * @param {File} params.file - 要上传的文件
 * @param {number} params.knowledgeId - 知识库ID
 * @param {Function} [onUploadProgress] - 上传进度回调
 * @param {Function} [onSuccess] - 上传成功回调
 * @param {Function} [onError] - 上传失败回调
 * @returns {Promise} 返回Promise对象
 */
export function uploadFile({ file, knowledgeId }, { onUploadProgress, onSuccess, onError } = {}) {
  if (!file) {
    const error = new Error('请选择要上传的文件')
    onError && onError(error)
    return Promise.reject(error)
  }
  
  if (!knowledgeId) {
    const error = new Error('请选择知识库')
    onError && onError(error)
    return Promise.reject(error)
  }
  
  const formData = new FormData()
  formData.append('file', file)
  formData.append('knowledgeId', knowledgeId)
  
  return request({
    url: '/api/rag/file/upload',
    method: 'post',
    data: formData,
    headers: {
      'Content-Type': 'multipart/form-data'
    },
    onUploadProgress: (progressEvent) => {
      if (onUploadProgress) {
        const percentCompleted = Math.round((progressEvent.loaded * 100) / progressEvent.total)
        onUploadProgress(percentCompleted, progressEvent)
      }
    }
  })
    .then(response => {
      onSuccess && onSuccess(response)
      return response
    })
    .catch(error => {
      onError && onError(error)
      return Promise.reject(error)
    })
}

// 批量文件上传并识别，给 FileUploader 组件使用
export function uploadAndRecognizeFile(formData) {
  return request({
    url: '/api/rag/file/upload',
    method: 'post',
    data: formData,
    headers: {
      'Content-Type': 'multipart/form-data'
    },
    showLoading: true
  })
}

/**
 * 获取文件列表
 * @param {Object} params - 查询参数
 * @param {number} params.knowledgeId - 知识库ID
 * @param {number} [params.pageNum=1] - 页码
 * @param {number} [params.pageSize=10] - 每页数量
 * @returns {Promise} 返回Promise对象
 */
export function getFileList({ knowledgeId, pageNum = 1, pageSize = 10 }) {
  if (!knowledgeId) {
    return Promise.reject(new Error('知识库ID不能为空'))
  }
  
  return request({
    url: '/api/rag/file/list',
    method: 'get',
    params: { knowledgeId },
    showLoading: true
  })
}

/**
 * 删除文件（支持批量）
 * @param {string|number|Array<string|number>} fileIds - 文件ID或ID数组
 * @returns {Promise} 返回Promise对象
 */
export function deleteFile(fileIds) {
  if (!fileIds) {
    return Promise.reject(new Error('文件ID不能为空'))
  }
  const ids = Array.isArray(fileIds) ? fileIds : [fileIds]

  return request({
    url: '/api/rag/file/delete',
    method: 'post',
    data: { fileIds: ids },
    showLoading: true
  })
}

/**
 * 预览文件
 * @param {string|number} fileId - 文件ID
 * @param {string} [disposition='inline'] - 文件处理方式，inline: 内联显示, attachment: 下载
 * @returns {string} 返回文件预览URL
 */
export function getFilePreviewUrl(fileId, disposition = 'inline') {
  if (!fileId) {
    throw new Error('文件ID不能为空')
  }
  
  const baseUrl = process.env.VUE_APP_BASE_API || ''
  return `${baseUrl}/api/rag/file/preview/${fileId}?disposition=${disposition}`
}

/**
 * 下载文件
 * @param {string|number} fileId - 文件ID
 * @param {string} [filename] - 下载文件名（可选）
 * @returns {Promise} 返回Promise对象
 */
export function downloadFile(fileId, filename) {
  if (!fileId) {
    return Promise.reject(new Error('文件ID不能为空'))
  }
  
  return request({
    url: `/api/rag/file/download/${fileId}`,
    method: 'get',
    responseType: 'blob',
    showLoading: true
  }).then(response => {
    // 创建blob对象
    const blob = new Blob([response])
    
    // 创建下载链接
    const downloadUrl = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = downloadUrl
    
    // 设置下载文件名
    if (filename) {
      link.setAttribute('download', filename)
    } else {
      // 从响应头中获取文件名
      const contentDisposition = response.headers['content-disposition']
      if (contentDisposition) {
        const fileNameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/)
        if (fileNameMatch && fileNameMatch[1]) {
          const fileName = fileNameMatch[1].replace(/['"]/g, '')
          link.setAttribute('download', decodeURIComponent(fileName))
        }
      }
    }
    
    // 触发下载
    document.body.appendChild(link)
    link.click()
    
    // 清理
    document.body.removeChild(link)
    window.URL.revokeObjectURL(downloadUrl)
    
    return response
  })
}