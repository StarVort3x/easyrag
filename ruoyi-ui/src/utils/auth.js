import Cookies from 'js-cookie'

const TokenKey = 'Admin-Token'
const TokenExpireKey = 'Token-Expire-Time'

export function getToken() {
  return Cookies.get(TokenKey)
}

export function setToken(token, rememberMe) {
  if (rememberMe) {
    // 记住我：7天过期
    return Cookies.set(TokenKey, token, { expires: 7 })
  } else {
    // 会话级存储，关闭浏览器后失效
    return Cookies.set(TokenKey, token)
  }
}

export function removeToken() {
  return Cookies.remove(TokenKey)
}

export function getTokenExpireTime() {
  return Cookies.get(TokenExpireKey)
}

export function setTokenExpireTime(expireTime) {
  return Cookies.set(TokenExpireKey, expireTime)
}

export function removeTokenExpireTime() {
  return Cookies.remove(TokenExpireKey)
}
