import { Clock3, Code2, FileText, Folder, Monitor } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import PageHeader from '../components/PageHeader'
import { getProjectTestcaseTree } from '../api/client'
import { getProjectById } from '../data/projectStore'

function ScopeFolder({ label, icon, screens, count, onClick }) {
  return (
    <button className="scope-folder-card" onClick={onClick}>
      <span className="scope-folder-icon">{icon}</span>
      <span className="scope-folder-copy">
        <b>{label}</b>
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
  const project = useMemo(() => getProjectById(projectId), [projectId])
  const [tree, setTree] = useState(null)
  const [backendReady, setBackendReady] = useState(true)

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

  return (
    <AppShell>
      <PageHeader index="02" title={project.name} subtitle="Chọn Web App hoặc API để mở các màn hình testcase." />

      <section className="card project-summary project-summary-clean">
        <div className="project-summary-main">
          <div className="big-folder"><Folder size={28} /></div>
          <div><h2>{project.name}</h2><p>{project.description}</p></div>
        </div>
        <div className="summary-stat"><FileText size={20} /><b>{totalTestcases}</b><span>testcase</span></div>
        <div className="summary-stat"><Clock3 size={20} /><b>{project.updated || 'vừa xong'}</b><span>cập nhật</span></div>
      </section>

      {!backendReady && (
        <div className="compact-warning">
          Backend chưa phản hồi. Bạn vẫn có thể mở Web App/API; hãy kiểm tra terminal backend nếu dữ liệu không tải được.
        </div>
      )}

      <section className="card explorer-root-card">
        <div className="explorer-section-head">
          <div><Folder size={22} /><h3>Thư mục dự án</h3></div>
          <span>{project.type === 'both' ? 2 : 1} thư mục</span>
        </div>

        <div className="scope-folder-grid">
          {project.type !== 'api' && (
            <ScopeFolder
              label="Web App"
              icon={<Monitor size={28} />}
              screens={(web.screens || []).length}
              count={webCount}
              onClick={() => navigate(`/project/${projectId}/scope/web`)}
            />
          )}
          {project.type !== 'web' && (
            <ScopeFolder
              label="API"
              icon={<Code2 size={28} />}
              screens={(api.screens || []).length}
              count={apiCount}
              onClick={() => navigate(`/project/${projectId}/scope/api`)}
            />
          )}
        </div>
      </section>
    </AppShell>
  )
}
