// API 封装：统一注入当前登录用户 X-User
const USER_KEY = 'planning_user'

export function getUsername() {
  return localStorage.getItem(USER_KEY) || ''
}
export function setUsername(u) {
  localStorage.setItem(USER_KEY, u)
}
export function clearUsername() {
  localStorage.removeItem(USER_KEY)
}

async function request(url, options = {}) {
  const headers = { 'X-User': getUsername(), ...(options.headers || {}) }
  if (options.body && !(options.body instanceof FormData)) {
    headers['Content-Type'] = 'application/json'
  }
  const resp = await fetch(url, { ...options, headers })
  let data
  try { data = await resp.json() } catch { data = { detail: '响应解析失败' } }
  if (!resp.ok) throw new Error(data.detail || '请求失败')
  return data
}

export const api = {
  login: (username, password) =>
    fetch('/api/login', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    }).then(async r => {
      const d = await r.json()
      if (!r.ok) throw new Error(d.detail || '登录失败')
      return d
    }),
  meta: () => request('/api/meta/modules'),
  moduleMeta: (code) => request(`/api/meta/module/${code}`),
  list: (code, q, page, pageSize, fieldQueries) =>
    request(`/api/data/${code}?q=${encodeURIComponent(q || '')}&fields_q=${encodeURIComponent(JSON.stringify(fieldQueries || {}))}&page=${page}&page_size=${pageSize}`),
  create: (code, body) => request(`/api/data/${code}`, { method: 'POST', body: JSON.stringify(body) }),
  update: (code, id, body) => request(`/api/data/${code}/${id}`, { method: 'PUT', body: JSON.stringify(body) }),
  remove: (code, ids) => request(`/api/data/${code}?ids=${ids.join(',')}`, { method: 'DELETE' }),
  importPreview: async (code, file) => {
    const fd = new FormData(); fd.append('file', file)
    return request(`/api/import/${code}/preview`, { method: 'POST', body: fd })
  },
  importCommit: (code, rows) => request(`/api/import/${code}/commit`, {
    method: 'POST', body: JSON.stringify({ rows })
  }),
  download: (code, type) => {
    const url = `/api/${type}/${code}`
    return fetch(url, { headers: { 'X-User': getUsername() } }).then(r => r.blob())
  },
  resetSeed: () => request('/api/seed/reset', { method: 'POST' }),
  roles: () => request('/api/roles'),
}
