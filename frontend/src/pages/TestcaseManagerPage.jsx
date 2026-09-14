import { ChevronDown, ChevronRight, Folder, FolderTree, Monitor, Search, Server, TestTube2 } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import { useNavigate, useParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import Breadcrumb from '../components/Breadcrumb'
import PageHeader from '../components/PageHeader'
import { getProjectTestcaseTree } from '../api/client'
import { getProjectById } from '../data/projectStore'

export default function TestcaseManagerPage() {
  const navigate = useNavigate()
  const { projectId } = useParams()
  const project = getProjectById(projectId)
  const [tree, setTree] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [query, setQuery] = useState('')
  const [open, setOpen] = useState({})

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    getProjectTestcaseTree(projectId)
      .then(data => { if (!cancelled) setTree(data) })
      .catch(e => { if (!cancelled) setError(e.message || 'Không tải được cây testcase.') })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [projectId])

  const scopes = useMemo(() => {
    const q = query.trim().toLowerCase()
    return (tree?.scopes || []).map(scope => ({
      ...scope,
      screens: (scope.screens || []).map(screen => ({
        ...screen,
        testcases: (screen.testcases || []).filter(tc => !q || screen.screen.toLowerCase().includes(q) || tc.id.toLowerCase().includes(q) || tc.name.toLowerCase().includes(q)),
      })).filter(screen => !q || screen.screen.toLowerCase().includes(q) || screen.testcases.length),
    })).filter(scope => scope.screens.length)
  }, [tree, query])

  const toggle = key => setOpen(prev => ({ ...prev, [key]: !prev[key] }))

  return (
    <AppShell wide>
      <Breadcrumb items={['Trang chủ', project?.name || projectId, 'Quản lý testcase']} />
      <PageHeader index="05" title="Quản lý testcase" subtitle="Quản lý theo cây Dự án → Web App/API → Màn hình → Testcase." />

      <section className="card manager-root-card">
        <div className="manager-project-head">
          <div className="big-folder"><FolderTree size={28} /></div>
          <div><small>DỰ ÁN</small><h2>{project?.name || projectId}</h2></div>
          <div className="manager-stats"><span><b>{tree?.total_screens || 0}</b> màn hình</span><span><b>{tree?.total_testcases || 0}</b> testcase</span></div>
        </div>
        <label className="search-box manager-search"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Tìm màn hình, ID hoặc tên testcase..." /></label>
      </section>

      {error && <div className="pipeline-alert danger"><b>Không tải được dữ liệu</b><span>{error}</span></div>}
      {loading && <div className="workspace-loading"><span className="spinner"></span><b>Đang dựng cây testcase...</b></div>}

      {!loading && scopes.map(scope => (
        <section className="card manager-scope" key={scope.key}>
          <div className="manager-scope-head">
            <div>{scope.key === 'api' ? <Server size={24} /> : <Monitor size={24} />}<span><h3>{scope.label}</h3><p>{scope.screens.length} folder màn hình</p></span></div>
          </div>

          <div className="manager-screen-list">
            {scope.screens.map(screen => {
              const key = `${scope.key}:${screen.screen}`
              const expanded = !!open[key]
              return <div className="manager-screen" key={key}>
                <button className="manager-screen-head" onClick={() => toggle(key)}>
                  <span className="manager-chevron">{expanded ? <ChevronDown size={18} /> : <ChevronRight size={18} />}</span>
                  <span className="manager-screen-icon"><Folder size={20} /></span>
                  <span className="manager-screen-title"><b>{screen.screen}</b><small>{screen.count} testcase</small></span>
                </button>

                {expanded && <div className="manager-testcase-list">
                  {screen.testcases.map(tc => (
                    <button
                      key={tc.recordId}
                      className="manager-testcase-row"
                      onClick={() => navigate(`/project/${projectId}/folder/${tc.folderId}/testcases?scope=${scope.key}&screen=${encodeURIComponent(screen.screen)}`)}
                    >
                      <TestTube2 size={16} />
                      <span className="manager-tc-id">{tc.id}</span>
                      <span className="manager-tc-name">{tc.name}</span>
                      <span className="type-chip">{tc.type}</span>
                    </button>
                  ))}
                </div>}
              </div>
            })}
          </div>
        </section>
      ))}
    </AppShell>
  )
}
