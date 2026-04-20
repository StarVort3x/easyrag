import axios from 'axios'
import { Message, Loading } from 'element-ui'
import { getToken } from '@/utils/auth'

// 创建Axios实例
const service = axios.create({
  baseURL: process.env.VUE_APP_BASE_API,
  timeout: 120000, // 请求超时时间120秒（2分钟）
  headers: {
    'Content-Type': 'application/json;charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest'
  }
})

// 请求队列
let loadingInstance = null
let requestCount = 0

// 显示loading
const showLoading = () => {
  if (requestCount === 0) {
    loadingInstance = Loading.service({
      lock: true,
      text: '加载中...',
      spinner: 'el-icon-loading',
      background: 'rgba(0, 0, 0, 0.7)'
    })
  }
  requestCount++
}

// 隐藏loading
const hideLoading = () => {
  requestCount--
  if (requestCount <= 0) {
    loadingInstance && loadingInstance.close()
    requestCount = 0
  }
}

// 请求拦截器
service.interceptors.request.use(
  config => {
    // 显示loading
    if (config.showLoading !== false) {
      showLoading()
    }
    
    // 添加token
    const token = getToken()
    if (token) {
      config.headers['Authorization'] = `Bearer ${token}`
    }
    
    return config
  },
  error => {
    // 关闭loading
    hideLoading()
    Message.error('请求发送失败，请检查网络连接')
    return Promise.reject(error)
  }
)

// 响应拦截器
service.interceptors.response.use(
  response => {
    // 关闭loading
    hideLoading()
    
    const res = response.data
    
    // 如果是文件下载，直接返回
    if (response.config.responseType === 'blob') {
      return res
    }

    // 处理标准响应格式
    if (res && typeof res === 'object') {
      if (res.code === 200 || res.code === 0) {
        return res.data !== undefined ? res.data : res
      } else {
        // 业务错误处理
        const errorMsg = res.msg || res.message || '请求失败'
        if (res.code === 401) {
          // token过期处理
          Message.error('登录已过期，请重新登录')
          // 这里可以添加跳转到登录页面的逻辑
        } else if (res.code === 403) {
          Message.error('没有权限访问该资源')
        } else {
          Message.error(errorMsg)
        }
        return Promise.reject(new Error(errorMsg))
      }
    }
    
    return res
  },
  error => {
    // 关闭loading
    hideLoading()
    
    let errorMessage = '请求失败，请稍后重试'
    
    if (error.response) {
      // 请求已发出，服务器返回状态码不在2xx范围内
      const status = error.response.status
      switch (status) {
        case 400:
          errorMessage = '请求参数错误'
          break
        case 401:
          errorMessage = '未授权，请重新登录'
          // 这里可以添加跳转到登录页面的逻辑
          break
        case 403:
          errorMessage = '拒绝访问'
          break
        case 404:
          errorMessage = '请求地址不存在'
          break
        case 408:
          errorMessage = '请求超时'
          break
        case 500:
          errorMessage = '服务器内部错误'
          break
        case 501:
          errorMessage = '服务未实现'
          break
        case 502:
          errorMessage = '网关错误'
          break
        case 503:
          errorMessage = '服务不可用'
          break
        case 504:
          errorMessage = '网关超时'
          break
        case 505:
          errorMessage = 'HTTP版本不受支持'
          break
        default:
          errorMessage = `连接错误 ${status}`
      }
    } else if (error.message.includes('timeout')) {
      errorMessage = '请求超时，请检查网络连接'
    } else if (error.message === 'Network Error') {
      errorMessage = '网络连接错误，请检查网络设置'
    }
    
    Message.error(errorMessage)
    return Promise.reject(error)
  }
)

export default service