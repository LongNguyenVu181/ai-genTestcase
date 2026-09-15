import { Folder, Home, KeyRound, Plus, Search } from 'lucide-react'
import { NavLink, useLocation, useNavigate, useParams } from 'react-router-dom'
import Brand from './Brand'
import { useEffect, useMemo, useState } from 'react'
import { getProjects, PROJECTS_UPDATED_EVENT, startProjectSync } from '../data/projectStore'

export default function Sidebar() {
  const navigate = useNavigate()
  const location = useLocation()
  const { projectId } = useParams()
  const [query, setQuery] = useState('')
  const [allProjects, setAllProjects] = useState(() => getProjects())

  useEffect(() => {
    let cancelled = false
    // One non-blocking metadata sync per app window. The UI always uses local cache first.
    startProjectSync().then(() => { if (!cancelled) setAllProjects(getProjects()) })
    const refresh = () => setAllProjects(getProjects())
    window.addEventListener(PROJECTS_UPDATED_EVENT, refresh)
    window.addEventListener('storage', refresh)
    return () => {
      cancelled = true
      window.removeEventListener(PROJECTS_UPDATED_EVENT, refresh)
      window.removeEventListener('storage', refresh)
    }
  }, [])

  const projects = useMemo(
    () => allProjects.filter(p => p.name.toLowerCase().includes(query.toLowerCase())),
    [allProjects, query],
  )

  const openNewProject = () => {
    if (location.pathname === '/dashboard') {
      window.dispatchEvent(new CustomEvent('testpilot:new-project'))
      if (location.search) navigate('/dashboard', { replace: true })
      return
    }
    navigate('/dashboard?mode=new-project')
  }

  return (
    <aside className="sidebar">
      <Brand />
      <nav className="sidebar-nav">
        <NavLink to="/dashboard" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
          <Home size={20} /> <span>Bảng điều khiển</span>
        </NavLink>
        <NavLink to="/api-keys" className={({ isActive }) => `nav-item ${isActive ? 'active' : ''}`}>
          <KeyRound size={20} /> <span>Quản lý khóa API</span>
        </NavLink>
      </nav>

      <button className="btn btn-primary btn-full sidebar-new" type="button" onClick={openNewProject}>
        <Plus size={18} /> Dự án mới
      </button>

      <label className="search-box sidebar-search">
        <Search size={18} />
        <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Tìm kiếm dự án..." />
      </label>

      <div className="sidebar-section-title">DANH SÁCH DỰ ÁN</div>
      <div className="project-list">
        {projects.map(project => (
          <button
            key={project.id}
            className={`project-link ${projectId === project.id ? 'active' : ''}`}
            type="button"
            onClick={() => navigate(`/project/${project.id}`)}
          >
            <Folder size={18} /> <span>{project.name}</span>
          </button>
        ))}
      </div>
    </aside>
  )
}
