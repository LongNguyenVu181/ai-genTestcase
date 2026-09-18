import JSZip from 'jszip'
import mammoth from 'mammoth/mammoth.browser'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import * as XLSX from 'xlsx'
import YAML from 'yaml'

pdfjs.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/legacy/build/pdf.worker.mjs', import.meta.url).toString()

const CONFIG_KEY = 'testpilot.ai.config.v2'
const SECRET_KEY = 'testpilot.ai.secret.v2'
const PROJECTS_KEY = 'testpilot.projects.v1'
const JOBS_KEY = 'testpilot.local.jobs.v1'
const RUNS_KEY = 'testpilot.local.runs.v1'
const TESTCASES_KEY = 'testpilot.local.testcases.v1'
const COUNTER_KEY = 'testpilot.local.record-counter.v1'
const JOBS_UPDATED_EVENT = 'testpilot:workspace-updated'

export const DEFAULT_CONFIG = {
  provider: 'Qwen',
  model: 'qwen-max',
  baseUrl: 'https://ws-2vuxxf5tta2cjplh.ap-southeast-1.maas.aliyuncs.com/compatible-mode/v1',
  apiKey: '',
}

const read = (key, fallback) => {
  try {
    const value = JSON.parse(sessionStorage.getItem(key) || '')
    return value ?? fallback
  } catch {
    return fallback
  }
}

const write = (key, value) => sessionStorage.setItem(key, JSON.stringify(value))
const clone = value => structuredClone(value)
const now = () => Date.now()
const uid = () => globalThis.crypto?.randomUUID?.() || `${now()}-${Math.random().toString(36).slice(2)}`
const publish = () => window.dispatchEvent(new CustomEvent(JOBS_UPDATED_EVENT))
const normalizeBaseUrl = value => String(value || '').trim().replace(/\/+$/, '')

const readStreamedModelContent = async response => {
  if (!response.body) throw new Error('AI không trả stream nội dung.')
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  let content = ''

  const consumeLine = line => {
    const value = line.trim()
    if (!value.startsWith('data:')) return
    const data = value.slice(5).trim()
    if (!data || data === '[DONE]') return
    try {
      const chunk = JSON.parse(data)
      const choice = chunk?.choices?.[0]
      if (chunk?.error?.message) throw new Error(chunk.error.message)
      if (choice?.delta?.content) content += choice.delta.content
    } catch (error) {
      if (error instanceof SyntaxError) return
      throw error
    }
  }

  while (true) {
    const { value, done } = await reader.read()
    buffer += decoder.decode(value || new Uint8Array(), { stream: !done })
    const lines = buffer.split(/\r?\n/)
    buffer = lines.pop() || ''
    lines.forEach(consumeLine)
    if (done) break
  }
  if (buffer) consumeLine(buffer)
  if (!content.trim()) throw new Error('AI không trả nội dung phân tích.')
  return content
}

export const loadAiConfig = () => {
  try {
    return { ...DEFAULT_CONFIG, ...read(CONFIG_KEY, {}), apiKey: sessionStorage.getItem(SECRET_KEY) || '' }
  } catch {
    return { ...DEFAULT_CONFIG }
  }
}

export const saveAiConfig = config => {
  const { apiKey, ...safe } = config
  write(CONFIG_KEY, safe)
  if (apiKey) sessionStorage.setItem(SECRET_KEY, apiKey)
  else sessionStorage.removeItem(SECRET_KEY)
}

export const maskKey = raw => {
  const clean = String(raw || '').trim()
  if (!clean) return '••••••••••••'
  const prefix = clean.startsWith('sk-') ? 'sk-' : ''
  return `${prefix}••••••••••${clean.slice(-4)}`
}

const requestModel = async ({ config, prompt }) => {
  let response
  try {
    response = await fetch('/api/model', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        baseUrl: normalizeBaseUrl(config.baseUrl),
        apiKey: config.apiKey.trim(),
        model: config.model.trim(),
        prompt,
        stream: true,
      }),
    })
  } catch (error) {
    throw new Error(`Không gọi được AI proxy của Site: ${error.message}`)
  }
  if (!response.ok) {
    const data = await response.json().catch(() => ({}))
    throw new Error(data?.error?.message || data?.message || `AI trả HTTP ${response.status}`)
  }
  return readStreamedModelContent(response)
}

export const testAiConfig = async config => {
  const text = await requestModel({
    config,
    prompt: 'Trả lời JSON hợp lệ duy nhất: {"result":"OK"}.',
  })
  return { ok: true, model: config.model, base_url: config.baseUrl, result: text }
}

async function extractPdf(file) {
  const document = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise
  const pages = []
  for (let index = 1; index <= document.numPages; index += 1) {
    const page = await document.getPage(index)
    const content = await page.getTextContent()
    pages.push(`--- TRANG ${index} ---\n${content.items.map(item => item.str).join(' ')}`)
  }
  return pages.join('\n\n')
}

async function extractXlsx(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' })
  return workbook.SheetNames.map(name => `=== SHEET: ${name} ===\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`).join('\n\n')
}

export async function extractFileText(file) {
  const name = String(file?.name || '').toLowerCase()
  if (name.endsWith('.pdf')) return extractPdf(file)
  if (name.endsWith('.docx')) return (await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })).value
  if (name.endsWith('.xlsx') || name.endsWith('.xls')) return extractXlsx(file)
  if (name.endsWith('.doc')) throw new Error('File .doc chưa được trình duyệt hỗ trợ. Hãy chuyển sang DOCX, PDF hoặc TXT.')
  return file.text()
}

const parseModelJson = raw => {
  const fenced = String(raw).match(/```(?:json)?\s*([\s\S]*?)```/i)
  const text = (fenced?.[1] || raw).trim()
  try {
    return JSON.parse(text)
  } catch {
    const start = text.indexOf('{')
    const end = text.lastIndexOf('}')
    if (start >= 0 && end > start) return JSON.parse(text.slice(start, end + 1))
    throw new Error('AI trả Rule Matrix không phải JSON hợp lệ.')
  }
}

const clean = value => String(value || '').trim()
const title = value => clean(value) || 'Chưa xác định'
const webType = category => ({ UI: 'Giao diện', VALIDATION: 'Kiểm tra dữ liệu', EXCEPTION: 'Ngoại lệ', BUSINESS_FLOW: 'Luồng' }[category] || 'Chức năng')
const apiType = category => ({ AUTH: 'Auth', PERMISSION: 'Permission', VALIDATION: 'Validation', HAPPY_PATH: 'Happy Path', BUSINESS_RULE: 'Business Rule' }[category] || 'Business Rule')

const validate = testcase => {
  const errors = {}
  if (!clean(testcase.name)) errors.name = 'Thiếu tên testcase'
  if (!clean(testcase.expectedResult)) errors.expectedResult = 'Thiếu kết quả mong đợi'
  if (!(testcase.steps || []).length) errors.steps = 'Thiếu bước kiểm thử'
  return { valid: !Object.keys(errors).length, errors }
}

const mapRule = (rule, index, scope, fallbackScreen) => {
  const category = clean(rule.category).toUpperCase() || (scope === 'api' ? 'BUSINESS_RULE' : 'ACTION')
  const screen = title(rule.screen || rule.target || fallbackScreen)
  const objective = clean(rule.test_objective || rule.objective || rule.name)
  const condition = clean(rule.test_condition || rule.condition)
  const endpoint = clean(rule.endpoint || rule.target)
  const testcase = {
    id: `TC_${String(index + 1).padStart(3, '0')}`,
    name: title(rule.rule_name || rule.name || objective),
    preCondition: clean(rule.precondition || rule.pre_condition),
    importance: clean(rule.importance),
    steps: Array.isArray(rule.steps) && rule.steps.length
      ? rule.steps.map(clean).filter(Boolean)
      : [scope === 'api' ? `Gửi request tới ${endpoint || 'API mục tiêu'} với dữ liệu: ${condition || 'dữ liệu hợp lệ'}.` : `Mở ${screen}.`, objective || 'Thực hiện thao tác theo yêu cầu.', condition ? `Nhập/chọn điều kiện: ${condition}.` : 'Quan sát phản hồi của hệ thống.'],
    testData: clean(rule.test_data || rule.data || condition),
    expectedResult: clean(rule.expected_result || rule.expected || 'Hệ thống xử lý đúng theo yêu cầu.'),
    actualResult: '', run1: '', run2: '', run3: '', currentResult: '', note: '', errorCode: '',
    qcWriter: '', sprintWriter: '', qcExecutor: '', sprintExecutor: '', reviewer: '', reviewDate: '', reviewContent: '',
    needAuto: '', automated: '', smoke: '', regression: '', outdated: '', outdatedDate: '',
    type: scope === 'api' ? apiType(category) : webType(category),
    sourceRuleId: clean(rule.rule_id || rule.id || `LOCAL-${index + 1}`),
    sourceRequirement: clean(rule.source_requirement || rule.requirement || objective),
    featureGroup: clean(rule.feature_group).toUpperCase() || (scope === 'api' ? 'API' : category === 'VALIDATION' ? 'VALIDATE' : category === 'UI' ? 'UI' : 'FUNCTION'),
    featureName: clean(rule.feature_name || rule.feature || screen),
    category,
    screen,
  }
  testcase.validation = validate(testcase)
  return testcase
}

const buildPrompt = ({ scope, sources }) => `Bạn là Senior QA. Hãy phân tích tài liệu ${scope === 'api' ? 'API Design và BA' : 'BA/Web'} bên dưới và trả về JSON duy nhất theo schema:
{"rules":[{"rule_id":"R-001","rule_name":"...","screen":"...","category":"${scope === 'api' ? 'AUTH|PERMISSION|VALIDATION|HAPPY_PATH|BUSINESS_RULE' : 'UI|VALIDATION|ACTION|DATA_GRID|BUSINESS_FLOW|EXCEPTION'}","feature_group":"...","feature_name":"...","test_objective":"...","test_condition":"...","test_data":"...","expected_result":"...","source_requirement":"...","steps":["..."]}]}
Tạo testcase cụ thể, không bịa endpoint/field không có trong tài liệu.\n\n${sources}`

const updateJob = (jobId, patch) => {
  const jobs = read(JOBS_KEY, {})
  if (!jobs[jobId]) return
  jobs[jobId] = { ...jobs[jobId], ...clone(patch), updated_at: now() }
  write(JOBS_KEY, jobs)
  publish()
}

const allTestcases = () => read(TESTCASES_KEY, [])
const saveTestcases = testcases => write(TESTCASES_KEY, testcases)
const allocateRecord = () => {
  const next = Number(sessionStorage.getItem(COUNTER_KEY) || '1')
  sessionStorage.setItem(COUNTER_KEY, String(next + 1))
  return next
}

const persistRunCases = ({ runId, projectId, folderId, scope, testcases }) => {
  const retained = allTestcases().filter(item => !(item.projectId === projectId && item.folderId === folderId && item.scope === scope))
  const mapped = testcases.map((testcase, index) => ({
    ...testcase, recordId: allocateRecord(), runId, projectId, folderId, scope, sortOrder: index,
  }))
  saveTestcases([...retained, ...mapped])
  return mapped
}

const beginAnalysis = ({ scope, files, config, projectId, folderId }) => {
  const jobId = uid()
  const jobs = read(JOBS_KEY, {})
  jobs[jobId] = {
    job_id: jobId, kind: scope, project_id: projectId, folder_id: folderId,
    status: 'queued', stage: 'queued', progress: 0, message: 'Đang khởi tạo phân tích trong trình duyệt...',
    source_names: files.map(file => file.name), error: {}, created_at: now(), updated_at: now(), run_id: null,
  }
  write(JOBS_KEY, jobs)
  window.setTimeout(async () => {
    try {
      updateJob(jobId, { status: 'processing', stage: 'extracting', progress: 10, message: 'Đang đọc nội dung tài liệu...' })
      const texts = await Promise.all(files.map(async file => `=== SOURCE: ${file.name} ===\n${await extractFileText(file)}`))
      if (!texts.some(text => clean(text))) throw new Error('Không trích xuất được nội dung tài liệu.')
      updateJob(jobId, { stage: 'analyzing', progress: 35, message: 'AI đang phân tích yêu cầu và tạo Rule Matrix...' })
      const raw = await requestModel({ config, prompt: buildPrompt({ scope, sources: texts.join('\n\n') }) })
      const matrix = parseModelJson(raw)
      const rules = Array.isArray(matrix?.rules) ? matrix.rules : []
      if (!rules.length) throw new Error('AI không tạo được rule testcase hợp lệ.')
      updateJob(jobId, { stage: 'mapping', progress: 80, message: 'Đang chuyển Rule Matrix thành testcase...' })
      const fallbackScreen = scope === 'api' ? 'API' : 'Web App'
      const generated = rules.map((rule, index) => mapRule(rule, index, scope, fallbackScreen))
      const runId = uid()
      const runs = read(RUNS_KEY, {})
      runs[runId] = { run_id: runId, kind: scope, project_id: projectId, folder_id: folderId, created_at: now(), source_names: files.map(file => file.name), matrix, agent1_summary: { mode: 'browser-direct-ai', final_rules: generated.length } }
      write(RUNS_KEY, runs)
      persistRunCases({ runId, projectId, folderId, scope, testcases: generated })
      updateJob(jobId, { status: 'completed', stage: 'completed', progress: 100, run_id: runId, message: `Hoàn tất ${generated.length} testcase.`, error: {} })
    } catch (error) {
      updateJob(jobId, { status: 'failed', stage: 'failed', message: error.message || 'Phân tích thất bại.', error: { message: error.message || 'Phân tích thất bại.' } })
    }
  }, 0)
  return { ok: true, job_id: jobId, status: 'queued', message: jobs[jobId].message }
}

export const analyzeWeb = ({ files, config, projectId, folderId }) => beginAnalysis({ scope: 'web', files, config, projectId, folderId })
export const analyzeApi = ({ designFile, baFile, config, projectId, folderId }) => beginAnalysis({ scope: 'api', files: [designFile, baFile], config, projectId, folderId })
export const getAnalysisJob = jobId => Promise.resolve(clone(read(JOBS_KEY, {})[jobId] || null)).then(job => {
  if (!job) throw new Error('Không tìm thấy tác vụ phân tích.')
  return job
})

export const invalidateProjectCache = () => publish()
export const listProjectMetadata = () => Promise.resolve({ projects: read(PROJECTS_KEY, []) })
export const saveProjectMetadata = project => Promise.resolve({ ok: true, project })
export const syncProjectMetadata = projects => Promise.resolve({ ok: true, projects })

export const getRun = runId => {
  const run = read(RUNS_KEY, {})[runId]
  if (!run) return Promise.reject(new Error('Không tìm thấy phiên phân tích.'))
  return Promise.resolve({ ok: true, ...clone(run), testcases: allTestcases().filter(item => item.runId === runId).sort((a, b) => a.sortOrder - b.sortOrder) })
}

export const getFolderTestcases = ({ projectId, folderId, scope }) => Promise.resolve({
  ok: true, project_id: projectId, folder_id: folderId, scope,
  testcases: allTestcases().filter(item => item.projectId === projectId && item.folderId === folderId && item.scope === scope).sort((a, b) => a.sortOrder - b.sortOrder),
})

export const getScreenTestcases = ({ projectId, scope, screen }) => Promise.resolve({
  ok: true, project_id: projectId, scope, screen,
  testcases: allTestcases().filter(item => item.projectId === projectId && item.scope === scope && (!screen || item.screen === screen)).sort((a, b) => a.sortOrder - b.sortOrder),
})

export const getProjectTestcaseTree = projectId => {
  const grouped = new Map()
  allTestcases().filter(item => item.projectId === projectId).forEach(item => {
    const key = `${item.scope}::${item.screen || 'Chưa xác định màn hình'}`
    if (!grouped.has(key)) grouped.set(key, { screen: item.screen || 'Chưa xác định màn hình', count: 0, folder_ids: [], folderCounts: {} })
    const node = grouped.get(key)
    node.count += 1
    if (!node.folder_ids.includes(item.folderId)) node.folder_ids.push(item.folderId)
    node.folderCounts[item.folderId] = (node.folderCounts[item.folderId] || 0) + 1
  })
  const scopes = ['web', 'api'].map(key => ({ key, label: key === 'api' ? 'API' : 'Web App', screens: [...grouped.entries()].filter(([entry]) => entry.startsWith(`${key}::`)).map(([, item]) => item) }))
  return Promise.resolve({ ok: true, project_id: projectId, total_testcases: [...grouped.values()].reduce((sum, item) => sum + item.count, 0), total_screens: grouped.size, scopes })
}

const emptyRun = ({ projectId, folderId, scope }) => {
  const runId = `manual::${scope}::${projectId}::${folderId}`
  const runs = read(RUNS_KEY, {})
  if (!runs[runId]) {
    runs[runId] = { run_id: runId, kind: scope, project_id: projectId, folder_id: folderId, created_at: now(), source_names: [], matrix: {}, agent1_summary: { manual: true } }
    write(RUNS_KEY, runs)
  }
  return runId
}

const normalizeCase = testcase => {
  const result = { ...clone(testcase), steps: Array.isArray(testcase.steps) ? testcase.steps : [] }
  result.validation = validate(result)
  return result
}

export const createTestcase = ({ projectId, folderId, scope, runId, testcase }) => {
  const cases = allTestcases()
  const recordId = allocateRecord()
  const item = { ...normalizeCase(testcase), recordId, runId: runId || emptyRun({ projectId, folderId, scope }), projectId, folderId, scope, sortOrder: cases.filter(caseItem => caseItem.projectId === projectId && caseItem.folderId === folderId && caseItem.scope === scope).length }
  saveTestcases([...cases, item])
  publish()
  return Promise.resolve({ ok: true, testcase: clone(item) })
}

export const updateTestcase = ({ recordId, testcase }) => {
  let updated = null
  const cases = allTestcases().map(item => {
    if (item.recordId !== recordId) return item
    updated = { ...item, ...normalizeCase(testcase), recordId: item.recordId, runId: item.runId, projectId: item.projectId, folderId: item.folderId, scope: item.scope, sortOrder: item.sortOrder }
    return updated
  })
  if (!updated) return Promise.reject(new Error('Không tìm thấy testcase.'))
  saveTestcases(cases)
  publish()
  return Promise.resolve({ ok: true, testcase: clone(updated) })
}

export const deleteTestcase = recordId => {
  const removed = allTestcases().find(item => item.recordId === recordId)
  if (!removed) return Promise.reject(new Error('Không tìm thấy testcase.'))
  const cases = allTestcases().filter(item => item.recordId !== recordId)
  let order = 0
  cases.forEach(item => {
    if (item.projectId === removed.projectId && item.folderId === removed.folderId && item.scope === removed.scope) {
      order += 1
      item.sortOrder = order - 1
      item.id = `TC_${String(order).padStart(3, '0')}`
    }
  })
  saveTestcases(cases)
  publish()
  return Promise.resolve({ ok: true })
}

const exportRows = cases => cases.map((item, index) => ({
  ID: `TC_${String(index + 1).padStart(3, '0')}`,
  Name: item.name, PreConditions: item.preCondition, Importance: item.importance,
  Step: (item.steps || []).join('\n'), 'Data test': item.testData, 'Expected Result': item.expectedResult,
  Type: item.type, Screen: item.screen, 'Source requirement': item.sourceRequirement,
}))

const download = (filename, blob) => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

const xmindTopic = (title, children = []) => ({ id: uid(), title, ...(children.length ? { children: { attached: children } } : {}) })

export const downloadTestcaseExport = async ({ projectId, folderId, scope, screen, type }) => {
  const cases = allTestcases().filter(item => item.projectId === projectId && item.scope === scope && (!folderId || item.folderId === folderId) && (!screen || item.screen === screen))
  if (!cases.length) throw new Error('Chưa có testcase để xuất.')
  const basename = (screen || folderId || 'TestCases').replace(/[^\w.-]+/g, '_')
  if (type === 'excel') {
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(exportRows(cases)), 'Test Cases')
    download(`${basename}.xlsx`, new Blob([XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
    return
  }
  const zip = new JSZip()
  const grouped = new Map()
  cases.forEach(item => {
    const key = item.featureName || item.screen || 'Testcases'
    if (!grouped.has(key)) grouped.set(key, [])
    grouped.get(key).push(xmindTopic(item.name, (item.steps || []).map(step => xmindTopic(step))))
  })
  zip.file('content.json', JSON.stringify([{ id: uid(), rootTopic: xmindTopic(screen || 'Test Cases', [...grouped.entries()].map(([name, children]) => xmindTopic(name, children))) }]))
  zip.file('metadata.json', JSON.stringify({ name: basename, creator: { name: 'TestPilot AI' }, createdTime: now(), modifiedTime: now() }))
  zip.file('manifest.json', JSON.stringify({ fileEntries: { 'content.json': {}, 'metadata.json': {} } }))
  download(`${basename}.xmind`, await zip.generateAsync({ type: 'blob' }))
}
