import { apiFolders as seedApiFolders, initialProjects, webFolders as seedWebFolders } from './mockData'

const STORAGE_KEY = 'testpilot.projects.v1'
export const PROJECTS_UPDATED_EVENT = 'testpilot:projects-updated'

const tones = ['blue', 'green', 'orange', 'purple']

const slugify = value => String(value || '')
  .trim()
  .toLowerCase()
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/đ/g, 'd')
  .replace(/[^a-z0-9]+/g, '-')
  .replace(/^-+|-+$/g, '') || 'du-an-moi'

const cloneFolders = items => items.map(item => ({ ...item }))

const seedProjects = initialProjects.map((project, index) => ({
  ...project,
  description: project.id === 'website-tmdt'
    ? 'Dự án kiểm thử cho website thương mại điện tử.'
    : 'Dự án kiểm thử được tạo sẵn để minh họa giao diện.',
  type: project.id === 'api-thanh-toan' ? 'api' : 'both',
  webFolders: project.id === 'api-thanh-toan' ? [] : cloneFolders(seedWebFolders),
  apiFolders: cloneFolders(seedApiFolders),
  createdAt: Date.now() - (index + 1) * 86400000,
}))

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

const emitUpdated = () => {
  window.dispatchEvent(new CustomEvent(PROJECTS_UPDATED_EVENT))
}

const writeProjects = projects => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(projects))
  emitUpdated()
  return projects
}

export const getProjects = () => {
  const stored = safeRead()
  if (stored?.length) return stored
  writeProjects(seedProjects)
  return seedProjects
}

export const getProjectById = projectId => {
  const project = getProjects().find(item => item.id === projectId)
  if (project) return project

  const seed = seedProjects.find(item => item.id === projectId)
  return seed || null
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
  .map((name, index) => ({
    id: slugify(name),
    name,
    count: 0,
    updated: 'vừa xong',
    tone: tones[index % tones.length],
  }))

export const createProject = ({ name, description = '', type = 'both', webFolders = [], apiFolders = [] }) => {
  const projects = getProjects()
  const id = makeUniqueId(name, projects)
  const now = Date.now()

  const project = {
    id,
    name: String(name || '').trim(),
    description: String(description || '').trim() || 'Dự án kiểm thử mới trên TestPilot AI.',
    type,
    webFolders: type === 'api' ? [] : folderObjectsFromNames(webFolders),
    apiFolders: type === 'web' ? [] : folderObjectsFromNames(apiFolders),
    updated: 'vừa xong',
    createdAt: now,
  }

  writeProjects([project, ...projects.filter(item => item.id !== id)])
  return project
}

export const addFolderToProject = (projectId, scope, folderName) => {
  const name = String(folderName || '').trim() || 'Thư mục mới'
  const projects = getProjects()
  const projectIndex = projects.findIndex(item => item.id === projectId)
  if (projectIndex < 0) return null

  const project = { ...projects[projectIndex] }
  const key = scope === 'api' ? 'apiFolders' : 'webFolders'
  const current = Array.isArray(project[key]) ? [...project[key]] : []
  let id = slugify(name)
  let idx = 2
  while (current.some(item => item.id === id)) {
    id = `${slugify(name)}-${idx}`
    idx += 1
  }

  const folder = {
    id,
    name,
    count: 0,
    updated: 'vừa xong',
    tone: tones[current.length % tones.length],
  }

  project[key] = [...current, folder]
  project.updated = 'vừa xong'
  projects[projectIndex] = project
  writeProjects(projects)
  return folder
}
