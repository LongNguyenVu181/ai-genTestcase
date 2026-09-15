import { listProjectMetadata, saveProjectMetadata, syncProjectMetadata } from '../api/client'

const STORAGE_KEY = 'testpilot.projects.v1'
const MIGRATION_KEY = 'testpilot.projects.backend-migrated.v1'
export const PROJECTS_UPDATED_EVENT = 'testpilot:projects-updated'

const tones = ['blue', 'green', 'orange', 'purple']
let syncPromise = null
let lastSyncAt = 0

const slugify = value => String(value || '')
  .trim()
  .toLowerCase()
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/đ/g, 'd')
  .replace(/[^a-z0-9]+/g, '-')
  .replace(/^-+|-+$/g, '') || 'du-an-moi'

const scopesFromType = type => type === 'api' ? ['api'] : type === 'web' ? ['web'] : ['web', 'api']


let memoryProjects = null

const safeRead = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return null
    const parsed = JSON.parse(raw)
    return Array.isArray(parsed) ? parsed : null
  } catch {
    return null
  }
}

const normalizeProject = project => {
  const type = ['web', 'api', 'both'].includes(project?.type) ? project.type : 'both'
  const enabledScopes = Array.isArray(project?.enabledScopes) && project.enabledScopes.length
    ? [...new Set(project.enabledScopes.filter(x => x === 'web' || x === 'api'))]
    : scopesFromType(type)
  return {
    ...project,
    id: String(project?.id || slugify(project?.name)),
    name: String(project?.name || 'Dự án'),
    description: String(project?.description || ''),
    type,
    enabledScopes,
    webFolders: Array.isArray(project?.webFolders) ? project.webFolders : [],
    apiFolders: Array.isArray(project?.apiFolders) ? project.apiFolders : [],
    customFolders: Array.isArray(project?.customFolders) ? project.customFolders : [],
    createdAt: Number(project?.createdAt || Date.now()),
    updatedAt: Number(project?.updatedAt || project?.createdAt || Date.now()),
  }
}

const emitUpdated = () => window.dispatchEvent(new CustomEvent(PROJECTS_UPDATED_EVENT))

const writeProjects = (projects, { emit = true } = {}) => {
  const normalized = projects.map(normalizeProject)
  memoryProjects = normalized
  localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized))
  if (emit) emitUpdated()
  return normalized
}

const persistProject = project => {
  saveProjectMetadata(normalizeProject(project)).catch(() => {})
}

export const getProjects = () => {
  if (Array.isArray(memoryProjects)) return memoryProjects
  const stored = safeRead()
  if (stored?.length) {
    memoryProjects = stored.map(normalizeProject)
    return memoryProjects
  }
  memoryProjects = []
  return memoryProjects
}

export const getProjectById = projectId => getProjects().find(item => item.id === projectId) || null

export const startProjectSync = ({ force = false } = {}) => {
  if (!force && Date.now() - lastSyncAt < 30000) return Promise.resolve(getProjects())
  if (syncPromise) return syncPromise
  syncPromise = (async () => {
    const local = getProjects()
    try {
      const response = await listProjectMetadata()
      const remote = Array.isArray(response?.projects) ? response.projects.map(normalizeProject) : []
      const merged = new Map()
      for (const project of [...remote, ...local]) {
        const current = merged.get(project.id)
        if (!current || Number(project.updatedAt || 0) >= Number(current.updatedAt || 0)) merged.set(project.id, project)
      }
      const projects = writeProjects([...merged.values()].sort((a, b) => Number(b.updatedAt) - Number(a.updatedAt)))

      // Migration only: older builds stored projects only in localStorage. Push only local-only
      // projects once, instead of POSTing the whole project list on every page load/reload.
      const migrated = localStorage.getItem(MIGRATION_KEY) === '1'
      if (!migrated && local.length) {
        const remoteIds = new Set(remote.map(project => project.id))
        const localOnly = local.filter(project => !remoteIds.has(project.id))
        if (localOnly.length) await syncProjectMetadata(localOnly).catch(() => null)
        localStorage.setItem(MIGRATION_KEY, '1')
      }

      lastSyncAt = Date.now()
      return projects
    } catch {
      lastSyncAt = Date.now()
      return local
    }
  })().finally(() => { syncPromise = null })
  return syncPromise
}

const makeUniqueId = (name, projects) => {
  const base = slugify(name)
  if (!projects.some(item => item.id === base)) return base
  let idx = 2
  while (projects.some(item => item.id === `${base}-${idx}`)) idx += 1
  return `${base}-${idx}`
}

const folderObjectsFromNames = (names = []) => names
  .map(name => String(name || '').trim())
  .filter(Boolean)
  .map((name, index) => ({ id: slugify(name), name, count: 0, updated: 'vừa xong', tone: tones[index % tones.length] }))

export const createProject = ({ name, description = '', type = 'both', webFolders = [], apiFolders = [] }) => {
  const projects = getProjects()
  const id = makeUniqueId(name, projects)
  const now = Date.now()
  const project = normalizeProject({
    id,
    name: String(name || '').trim(),
    description: String(description || '').trim() || 'Dự án kiểm thử mới trên TestPilot AI.',
    type,
    enabledScopes: scopesFromType(type),
    webFolders: type === 'api' ? [] : folderObjectsFromNames(webFolders),
    apiFolders: type === 'web' ? [] : folderObjectsFromNames(apiFolders),
    customFolders: [],
    updated: 'vừa xong',
    createdAt: now,
    updatedAt: now,
  })
  writeProjects([project, ...projects.filter(item => item.id !== id)])
  persistProject(project)
  return project
}

export const enableProjectScope = (projectId, scope) => {
  const normalizedScope = scope === 'api' ? 'api' : 'web'
  const projects = getProjects()
  const projectIndex = projects.findIndex(item => item.id === projectId)
  if (projectIndex < 0) return null
  const project = normalizeProject({ ...projects[projectIndex] })
  project.enabledScopes = [...new Set([...(project.enabledScopes || []), normalizedScope])]
  project.type = project.enabledScopes.length > 1 ? 'both' : normalizedScope
  project.updated = 'vừa xong'
  project.updatedAt = Date.now()
  projects[projectIndex] = project
  writeProjects(projects)
  persistProject(project)
  return project
}

export const addProjectFolder = (projectId, scope, folderName) => {
  const name = String(folderName || '').trim()
  if (!name) return null
  const projects = getProjects()
  const projectIndex = projects.findIndex(item => item.id === projectId)
  if (projectIndex < 0) return null
  const project = normalizeProject({ ...projects[projectIndex] })
  const current = [...project.customFolders]
  const normalizedScope = scope === 'api' ? 'api' : 'web'
  let id = slugify(name)
  let idx = 2
  while (current.some(item => item.id === id)) id = `${slugify(name)}-${idx++}`
  const folder = { id, name, scope: normalizedScope, createdAt: Date.now() }
  project.customFolders = [...current, folder]
  project.enabledScopes = [...new Set([...(project.enabledScopes || []), normalizedScope])]
  project.type = project.enabledScopes.length > 1 ? 'both' : normalizedScope
  project.updated = 'vừa xong'
  project.updatedAt = Date.now()
  projects[projectIndex] = project
  writeProjects(projects)
  persistProject(project)
  return folder
}

