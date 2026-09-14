import { ArrowLeft, FilePlus2, Folder, Monitor, Search, Server } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import { getProjectTestcaseTree } from '../api/client'
import { getProjectById } from '../data/projectStore'

export default function ScopeExplorerPage() {
  const navigate = useNavigate()
  const { projectId, scope: scopeParam } = useParams()
  const scope = scopeParam === 'api' ? 'api' : 'web'
  const project = getProjectById(projectId)
  const [tree, setTree] = useState(null)
  const [query, setQuery] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    setError('')
    getProjectTestcaseTree(projectId)
      .then(data => {
        if (cancelled) return
        setTree(data)
        const node = (data.scopes || []).find(item => item.key === scope)
        if (!node || !(node.screens || []).length) {
          navigate(`/project/${projectId}/scope/${scope}/upload`, { replace: true })
        }
      })
      .catch(e => {
        if (!cancelled) setError(e.message || 'Không tải được dữ liệu.')
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => { cancelled = true }
  }, [projectId, scope, navigate])

  const node = (tree?.scopes || []).find(item => item.key === scope) || { screens: [] }
  const screens = useMemo(() => {
    const q = query.trim().toLowerCase()
    return (node.screens || []).filter(item => !q || item.screen.toLowerCase().includes(q))
  }, [node.screens, query])

  const total = (node.screens || []).reduce((sum, item) => sum + Number(item.count || 0), 0)
  const label = scope === 'api' ? 'API' : 'Web App'

  return (
    <AppShell wide>
      <div className="clean-page-head">
        <button className="icon-text-button" onClick={() => navigate(`/project/${projectId}`)}>
          <ArrowLeft size={18} /> Quay lại
        </button>
        <div className="clean-title">
          <span className="clean-title-icon">{scope === 'api' ? <Server size={26} /> : <Monitor size={26} />}</span>
          <div><h1>{label}</h1><p>{project?.name || projectId} · {node.screens?.length || 0} màn hình · {total} testcase</p></div>
        </div>
        <button className="btn btn-primary btn-balanced" onClick={() => navigate(`/project/${projectId}/scope/${scope}/upload`)}>
          <FilePlus2 size={18} /> Phân tích tài liệu mới
        </button>
      </div>

      <section className="card explorer-toolbar">
        <label className="search-box explorer-search">
          <Search size={19} />
          <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Tìm màn hình..." />
        </label>
      </section>

      {error && <div className="pipeline-alert danger"><b>Không tải được dữ liệu</b><span>{error}</span></div>}
      {loading && <div className="workspace-loading"><span className="spinner"></span><b>Đang tải...</b></div>}

      {!loading && !error && (
        <section className="screen-folder-grid">
          {screens.map(screen => (
            <button
              className="screen-folder-card"
              key={screen.screen}
              onClick={() => navigate(`/project/${projectId}/scope/${scope}/testcases?screen=${encodeURIComponent(screen.screen)}`)}
            >
              <span className="screen-folder-icon"><Folder size={27} /></span>
              <span className="screen-folder-copy">
                <b>{screen.screen}</b>
                <small>{screen.count} testcase</small>
              </span>
              <span className="screen-folder-open">Mở</span>
            </button>
          ))}
        </section>
      )}
    </AppShell>
  )
}
