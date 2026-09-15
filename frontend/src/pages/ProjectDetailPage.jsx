import { Clock3, Code2, FileText, Folder, Monitor, Plus } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import Modal from '../components/Modal'
import PageHeader from '../components/PageHeader'
import { getProjectTestcaseTree } from '../api/client'
import { addProjectFolder, enableProjectScope, getProjectById, PROJECTS_UPDATED_EVENT } from '../data/projectStore'

function ScopeFolder({ label, icon, screens, count, onClick, meta }) {
  return (
    <button className="scope-folder-card" onClick={onClick}>
      <span className="scope-folder-icon">{icon}</span>
      <span className="scope-folder-copy">
        <b>{label}</b>
        {meta && <small>{meta}</small>}
        <small>{screens} màn hình</small>
        <small>{count} testcase</small>
      </span>
      <span className="scope-folder-open">Mở</span>
    </button>
  )
}

export default function ProjectDetailPage() {
  const navigate = useNavigate()
  const { projectId = 'website-tmdt' } = useParams()
  const [project, setProject] = useState(() => getProjectById(projectId))
  const [tree, setTree] = useState(null)
  const [backendReady, setBackendReady] = useState(true)
  const [folderModalOpen, setFolderModalOpen] = useState(false)
  const [folderName, setFolderName] = useState('')
  const [folderScope, setFolderScope] = useState('api')

  useEffect(() => {
    const refresh = () => setProject(getProjectById(projectId))
    refresh()
    window.addEventListener(PROJECTS_UPDATED_EVENT, refresh)
    window.addEventListener('storage', refresh)
    return () => {
      window.removeEventListener(PROJECTS_UPDATED_EVENT, refresh)
      window.removeEventListener('storage', refresh)
    }
  }, [projectId])

  useEffect(() => {
    let cancelled = false
    getProjectTestcaseTree(projectId)
      .then(data => {
        if (!cancelled) {
          setTree(data)
          setBackendReady(true)
        }
      })
      .catch(() => {
        if (!cancelled) setBackendReady(false)
      })
    return () => { cancelled = true }
  }, [projectId])

  const enabledScopes = useMemo(() => {
    if (!project) return []
    if (Array.isArray(project.enabledScopes) && project.enabledScopes.length) return project.enabledScopes
    return project.type === 'api' ? ['api'] : project.type === 'web' ? ['web'] : ['web', 'api']
  }, [project])

  if (!project) {
    return (
      <AppShell>
        <PageHeader index="02" title="Không tìm thấy dự án" subtitle="" />
        <button className="btn btn-primary" onClick={() => navigate('/dashboard')}>Quay về Bảng điều khiển</button>
      </AppShell>
    )
  }

  const scopeData = key => (tree?.scopes || []).find(item => item.key === key) || { screens: [] }
  const web = scopeData('web')
  const api = scopeData('api')
  const webCount = (web.screens || []).reduce((sum, x) => sum + Number(x.count || 0), 0)
  const apiCount = (api.screens || []).reduce((sum, x) => sum + Number(x.count || 0), 0)
  const totalTestcases = Number(tree?.total_testcases || 0)
  const customFolders = Array.isArray(project.customFolders) ? project.customFolders : []

  const customFolderStats = folder => {
    const node = scopeData(folder.scope)
    let count = 0
    let screens = 0
    for (const screen of node.screens || []) {
      const folderCount = Number(screen.folderCounts?.[folder.id] || 0)
      if (folderCount > 0) {
        screens += 1
        count += folderCount
      }
    }
    return { screens, count }
  }

  const openAddScope = () => {
    setFolderName('')
    setFolderScope(enabledScopes.includes('api') && !enabledScopes.includes('web') ? 'web' : 'api')
    setFolderModalOpen(true)
  }

  const selectedScopeExists = enabledScopes.includes(folderScope)

  const confirmAdd = () => {
    if (!selectedScopeExists) {
      const updated = enableProjectScope(projectId, folderScope)
      if (!updated) return
      setProject(updated)
      setFolderModalOpen(false)
      // Open upload immediately so adding API feels like a direct user action, not a filesystem task.
      navigate(`/project/${projectId}/scope/${folderScope}/upload`)
      return
    }

    const cleanName = folderName.trim()
    if (!cleanName) return
    const folder = addProjectFolder(projectId, folderScope, cleanName)
    if (!folder) return
    setProject(getProjectById(projectId))
    setFolderModalOpen(false)
    setFolderName('')
  }

  const openCustomFolder = folder => {
    const stats = customFolderStats(folder)
    const encodedScope = encodeURIComponent(folder.scope)
    if (stats.count > 0) navigate(`/project/${projectId}/folder/${folder.id}/testcases?scope=${encodedScope}`)
    else navigate(`/project/${projectId}/folder/${folder.id}/upload?scope=${encodedScope}`)
  }

  return (
    <AppShell>
      <PageHeader index="02" title={project.name} subtitle="Chọn Web App hoặc API để mở workspace testcase." />

      <section className="card project-summary project-summary-clean">
        <div className="project-summary-main">
          <div className="big-folder"><Folder size={28} /></div>
          <div><h2>{project.name}</h2><p>{project.description}</p></div>
        </div>
        <div className="summary-stat"><FileText size={20} /><b>{totalTestcases}</b><span>testcase</span></div>
        <div className="summary-stat"><Clock3 size={20} /><b>{project.updated || 'vừa xong'}</b><span>cập nhật</span></div>
      </section>

      {!backendReady && (
        <div className="compact-warning">Backend chưa phản hồi. Dữ liệu dự án vẫn được cache cục bộ và sẽ đồng bộ lại khi server sẵn sàng.</div>
      )}

      <section className="card explorer-root-card">
        <div className="explorer-section-head">
          <div><Folder size={22} /><h3>Thư mục dự án</h3></div>
          <div className="explorer-head-actions">
            <span>{enabledScopes.length + customFolders.length} workspace</span>
            <button className="btn btn-primary btn-balanced explorer-new-folder" type="button" onClick={openAddScope}>
              <Plus size={17} /> Thêm Web App / API
            </button>
          </div>
        </div>

        <div className="scope-folder-grid">
          {enabledScopes.includes('web') && (
            <ScopeFolder label="Web App" icon={<Monitor size={28} />} screens={(web.screens || []).length} count={webCount} onClick={() => navigate(`/project/${projectId}/scope/web`)} />
          )}
          {enabledScopes.includes('api') && (
            <ScopeFolder label="API" icon={<Code2 size={28} />} screens={(api.screens || []).length} count={apiCount} onClick={() => navigate(`/project/${projectId}/scope/api`)} />
          )}
          {customFolders.map(folder => {
            const stats = customFolderStats(folder)
            return (
              <ScopeFolder
                key={`${folder.scope}-${folder.id}`}
                label={folder.name}
                icon={folder.scope === 'api' ? <Code2 size={28} /> : <Folder size={28} />}
                meta={folder.scope === 'api' ? 'API workspace' : 'Web App workspace'}
                screens={stats.screens}
                count={stats.count}
                onClick={() => openCustomFolder(folder)}
              />
            )
          })}
        </div>
      </section>

      <Modal open={folderModalOpen} title="Bổ sung Web App / API" onClose={() => setFolderModalOpen(false)} width="600px">
        <div className="modal-body new-folder-form">
          <p className="scope-helper">Chọn loại muốn bổ sung. Nếu loại này chưa có trong dự án, Hệ thống sẽ tạo workspace và mở ngay màn tải tài liệu.</p>

          <div className="scope-choice-grid">
            <button type="button" className={`scope-choice ${folderScope === 'web' ? 'active' : ''}`} onClick={() => setFolderScope('web')}>
              <Monitor size={24} /><span><b>Web App</b><small>{enabledScopes.includes('web') ? 'Đã có · có thể tạo workspace riêng' : 'Bổ sung Web App vào dự án'}</small></span>
            </button>
            <button type="button" className={`scope-choice ${folderScope === 'api' ? 'active' : ''}`} onClick={() => setFolderScope('api')}>
              <Code2 size={24} /><span><b>API</b><small>{enabledScopes.includes('api') ? 'Đã có · có thể tạo workspace riêng' : 'Bổ sung API vào dự án'}</small></span>
            </button>
          </div>

          {selectedScopeExists && (
            <label>
              <span>Tên workspace riêng <em>*</em></span>
              <input className="input" autoFocus value={folderName} onChange={e => setFolderName(e.target.value)} placeholder={folderScope === 'api' ? 'Ví dụ: API Duyệt giao dịch' : 'Ví dụ: Màn hình phê duyệt'} />
              <small className="field-help">Nếu chỉ cần phân tích thêm tài liệu vào workspace {folderScope === 'api' ? 'API' : 'Web App'} hiện tại, hãy đóng cửa sổ và mở workspace đó.</small>
            </label>
          )}

          <div className="actions-end">
            <button className="btn btn-outline" type="button" onClick={() => setFolderModalOpen(false)}>Hủy</button>
            <button className="btn btn-primary" type="button" disabled={selectedScopeExists && !folderName.trim()} onClick={confirmAdd}>
              <Plus size={17} /> {selectedScopeExists ? 'Tạo workspace' : `Thêm ${folderScope === 'api' ? 'API' : 'Web App'}`}
            </button>
          </div>
        </div>
      </Modal>
    </AppShell>
  )
}
