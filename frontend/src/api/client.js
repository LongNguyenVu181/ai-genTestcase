const CONFIG_KEY = 'testpilot.ai.config.v2'
const SECRET_KEY = 'testpilot.ai.secret.v2'

export const DEFAULT_CONFIG = {
  provider: 'Qwen',
  model: 'qwen-max',
  baseUrl: 'https://ws-2vuxxf5tta2cjplh.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1',
  apiKey: '',
}

export const loadAiConfig = () => {
  try {
    return { ...DEFAULT_CONFIG, ...JSON.parse(localStorage.getItem(CONFIG_KEY) || '{}'), apiKey: sessionStorage.getItem(SECRET_KEY) || '' }
  } catch {
    return { ...DEFAULT_CONFIG }
  }
}

export const saveAiConfig = config => {
  const { apiKey, ...safe } = config
  localStorage.setItem(CONFIG_KEY, JSON.stringify(safe))
  if (apiKey) sessionStorage.setItem(SECRET_KEY, apiKey)
  else sessionStorage.removeItem(SECRET_KEY)
}

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
    const detail = payload?.detail
    const message = typeof detail === 'string'
      ? detail
      : detail?.message || payload?.message || (typeof payload === 'string' ? payload : `HTTP ${response.status}`)
    const err = new Error(message)
    err.payload = payload
    throw err
  }
  return payload
}

export const testAiConfig = config => fetch('/api/config/test', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    api_key: config.apiKey,
    base_url: config.baseUrl,
    model: config.model,
  }),
}).then(parseResponse)

export const analyzeWeb = async ({ files, config, projectId, folderId }) => {
  const form = new FormData()
  files.forEach(file => form.append('files', file))
  form.append('api_key', config.apiKey)
  form.append('base_url', config.baseUrl)
  form.append('model', config.model)
  form.append('project_id', projectId)
  form.append('folder_id', folderId)
  return fetch('/api/web/analyze', { method: 'POST', body: form }).then(parseResponse)
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
  return fetch('/api/api/analyze', { method: 'POST', body: form }).then(parseResponse)
}

export const getRun = runId => fetch(`/api/runs/${runId}`).then(parseResponse)

export const getFolderTestcases = ({ projectId, folderId, scope }) =>
  fetch(`/api/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(folderId)}/testcases?scope=${encodeURIComponent(scope)}`).then(parseResponse)

export const createTestcase = ({ projectId, folderId, scope, runId, testcase }) => fetch('/api/testcases', {
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

export const updateTestcase = ({ recordId, testcase }) => fetch(`/api/testcases/${recordId}`, {
  method: 'PUT',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ testcase }),
}).then(parseResponse)

export const deleteTestcase = recordId => fetch(`/api/testcases/${recordId}`, {
  method: 'DELETE',
}).then(parseResponse)

export const folderDownloadUrl = ({ projectId, folderId, scope, type }) =>
  `/api/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(folderId)}/${type}?scope=${encodeURIComponent(scope)}`

export const downloadUrl = (runId, type) => `/api/runs/${runId}/${type}`


export const getProjectTestcaseTree = projectId =>
  fetch(`/api/projects/${encodeURIComponent(projectId)}/testcase-tree`).then(parseResponse)


export const getScreenTestcases = ({ projectId, scope, screen }) =>
  fetch(`/api/projects/${encodeURIComponent(projectId)}/screen-testcases?scope=${encodeURIComponent(scope)}&screen=${encodeURIComponent(screen)}`).then(parseResponse)

export const screenDownloadUrl = ({ projectId, scope, screen, type }) =>
  `/api/projects/${encodeURIComponent(projectId)}/screen-${type}?scope=${encodeURIComponent(scope)}&screen=${encodeURIComponent(screen)}`
