const CONFIG_KEY = 'testpilot.ai.config.v2'
const SECRET_KEY = 'testpilot.ai.secret.v2'
const SESSION_ID_KEY = 'testpilot.browser-session.v1'

export const DEFAULT_CONFIG = {
  provider: 'Qwen',
  model: 'qwen-max',
  baseUrl: 'https://ws-2vuxxf5tta2cjplh.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1',
  apiKey: '',
}

export const loadAiConfig = () => {
  try {
    return { ...DEFAULT_CONFIG, ...JSON.parse(sessionStorage.getItem(CONFIG_KEY) || '{}'), apiKey: sessionStorage.getItem(SECRET_KEY) || '' }
  } catch {
    return { ...DEFAULT_CONFIG }
  }
}

export const saveAiConfig = config => {
  const { apiKey, ...safe } = config
  sessionStorage.setItem(CONFIG_KEY, JSON.stringify(safe))
  if (apiKey) sessionStorage.setItem(SECRET_KEY, apiKey)
  else sessionStorage.removeItem(SECRET_KEY)
}

export const getBrowserSessionId = () => {
  let sessionId = sessionStorage.getItem(SESSION_ID_KEY)
  if (!sessionId) {
    sessionId = globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`
    sessionStorage.setItem(SESSION_ID_KEY, sessionId)
  }
  return sessionId
}

const apiFetch = (input, init = {}) => {
  const headers = new Headers(init.headers)
  headers.set('X-TestPilot-Session', getBrowserSessionId())
  return fetch(input, { ...init, headers })
}

const withSessionQuery = path => `${path}${path.includes('?') ? '&' : '?'}session_id=${encodeURIComponent(getBrowserSessionId())}`

export const maskKey = raw => {
  const clean = String(raw || '').trim()
  if (!clean) return '••••••••••••'
  const prefix = clean.startsWith('sk-') ? 'sk-' : ''
  return `${prefix}••••••••••${clean.slice(-4)}`
}

const parseResponse = async response => {
  const type = response.headers.get('content-type') || ''
  const payload = type.includes('application/json') ? await response.json() : await response.text()
  if (!response.ok) {
    const isHtml = type.includes('text/html') || (typeof payload === 'string' && /<!doctype html|<html/i.test(payload))
    let message
    if (isHtml && response.status === 524) {
      message = 'Kết nối tới server bị timeout. Tác vụ AI có thể vẫn đang chạy; hãy kiểm tra trạng thái sau ít phút.'
    } else if (isHtml && response.status >= 500) {
      message = 'Không kết nối ổn định tới server. Vui lòng thử lại sau ít phút.'
    } else {
      const detail = payload?.detail
      message = typeof detail === 'string'
        ? detail
        : detail?.message || payload?.message || (typeof payload === 'string' ? payload : `HTTP ${response.status}`)
    }
    const err = new Error(message)
    err.payload = payload
    err.status = response.status
    throw err
  }
  return payload
}

export const testAiConfig = config => apiFetch('/api/config/test', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    api_key: config.apiKey,
    base_url: config.baseUrl,
    model: config.model,
  }),
}).then(parseResponse)


const responseCache = new Map()
const inflight = new Map()
const CACHE_TTL_MS = 20000

const cachedJson = (key, factory, ttl = CACHE_TTL_MS) => {
  const now = Date.now()
  const hit = responseCache.get(key)
  if (hit && now - hit.at < ttl) return Promise.resolve(hit.value)
  if (inflight.has(key)) return inflight.get(key)
  const promise = factory()
    .then(value => {
      responseCache.set(key, { value, at: Date.now() })
      inflight.delete(key)
      return value
    })
    .catch(error => {
      inflight.delete(key)
      throw error
    })
  inflight.set(key, promise)
  return promise
}

export const invalidateProjectCache = projectId => {
  const prefix = `project:${projectId}:`
  for (const key of responseCache.keys()) {
    if (key.startsWith(prefix)) responseCache.delete(key)
  }
}

export const listProjectMetadata = () =>
  apiFetch('/api/projects').then(parseResponse)

export const saveProjectMetadata = project =>
  apiFetch(`/api/projects/${encodeURIComponent(project.id)}/metadata`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(project),
  }).then(parseResponse)

export const syncProjectMetadata = projects =>
  apiFetch('/api/projects/sync', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ projects }),
  }).then(parseResponse)

export const analyzeWeb = async ({ files, config, projectId, folderId }) => {
  const form = new FormData()
  files.forEach(file => form.append('files', file))
  form.append('api_key', config.apiKey)
  form.append('base_url', config.baseUrl)
  form.append('model', config.model)
  form.append('project_id', projectId)
  form.append('folder_id', folderId)
  return apiFetch('/api/web/analyze', { method: 'POST', body: form }).then(parseResponse)
}

export const analyzeApi = async ({ designFile, baFile, config, projectId, folderId }) => {
  const form = new FormData()
  form.append('design_file', designFile)
  form.append('ba_file', baFile)
  form.append('api_key', config.apiKey)
  form.append('base_url', config.baseUrl)
  form.append('model', config.model)
  form.append('project_id', projectId)
  form.append('folder_id', folderId)
  return apiFetch('/api/api/analyze', { method: 'POST', body: form }).then(parseResponse)
}

export const getAnalysisJob = jobId =>
  apiFetch(`/api/jobs/${encodeURIComponent(jobId)}`, { cache: 'no-store' }).then(parseResponse)

export const getRun = runId => apiFetch(`/api/runs/${runId}`).then(parseResponse)

export const getFolderTestcases = ({ projectId, folderId, scope }) =>
  cachedJson(`project:${projectId}:folder:${scope}:${folderId}`, () =>
    apiFetch(`/api/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(folderId)}/testcases?scope=${encodeURIComponent(scope)}`).then(parseResponse), 12000)

export const createTestcase = ({ projectId, folderId, scope, runId, testcase }) => apiFetch('/api/testcases', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    project_id: projectId,
    folder_id: folderId,
    scope,
    run_id: runId || null,
    testcase,
  }),
}).then(parseResponse)

export const updateTestcase = ({ recordId, testcase }) => apiFetch(`/api/testcases/${recordId}`, {
  method: 'PUT',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ testcase }),
}).then(parseResponse)

export const deleteTestcase = recordId => apiFetch(`/api/testcases/${recordId}`, {
  method: 'DELETE',
}).then(parseResponse)

export const folderDownloadUrl = ({ projectId, folderId, scope, type }) =>
  withSessionQuery(`/api/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(folderId)}/${type}?scope=${encodeURIComponent(scope)}`)

export const downloadUrl = (runId, type) => withSessionQuery(`/api/runs/${runId}/${type}`)


export const getProjectTestcaseTree = projectId =>
  cachedJson(`project:${projectId}:tree`, () =>
    apiFetch(`/api/projects/${encodeURIComponent(projectId)}/testcase-tree`).then(parseResponse), 20000)


export const getScreenTestcases = ({ projectId, scope, screen }) =>
  cachedJson(`project:${projectId}:screen:${scope}:${screen}`, () =>
    apiFetch(`/api/projects/${encodeURIComponent(projectId)}/screen-testcases?scope=${encodeURIComponent(scope)}&screen=${encodeURIComponent(screen)}`).then(parseResponse), 12000)

export const screenDownloadUrl = ({ projectId, scope, screen, type }) =>
  withSessionQuery(`/api/projects/${encodeURIComponent(projectId)}/screen-${type}?scope=${encodeURIComponent(scope)}&screen=${encodeURIComponent(screen)}`)
