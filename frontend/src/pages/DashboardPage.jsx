import { ArrowRight, FolderOpen } from 'lucide-react'
import { useLocation, useNavigate } from 'react-router-dom'
import { useEffect, useRef, useState } from 'react'
import AppShell from '../components/AppShell'
import Breadcrumb from '../components/Breadcrumb'
import PageHeader from '../components/PageHeader'
import { createProject, getProjects, PROJECTS_UPDATED_EVENT } from '../data/projectStore'

export default function DashboardPage() {
  const navigate = useNavigate()
  const location = useLocation()
  const nameInputRef = useRef(null)
  const [form, setForm] = useState({ name: '', description: '', type: '' })
  const [projects, setProjects] = useState(() => getProjects())
  const [submitError, setSubmitError] = useState('')

  useEffect(() => {
    const refresh = () => setProjects(getProjects())
    const focusNewProject = () => {
      setSubmitError('')
      requestAnimationFrame(() => {
        nameInputRef.current?.focus({ preventScroll: true })
        nameInputRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' })
      })
    }
    window.addEventListener(PROJECTS_UPDATED_EVENT, refresh)
    window.addEventListener('storage', refresh)
    window.addEventListener('testpilot:new-project', focusNewProject)
    return () => {
      window.removeEventListener(PROJECTS_UPDATED_EVENT, refresh)
      window.removeEventListener('storage', refresh)
      window.removeEventListener('testpilot:new-project', focusNewProject)
    }
  }, [])

  useEffect(() => {
    if (new URLSearchParams(location.search).get('mode') === 'new-project') {
      requestAnimationFrame(() => nameInputRef.current?.focus({ preventScroll: true }))
    }
  }, [location.search])

  const submit = e => {
    e.preventDefault()
    setSubmitError('')
    if (!form.name.trim()) {
      setSubmitError('Vui lòng nhập tên dự án.')
      nameInputRef.current?.focus()
      return
    }
    if (!form.type) {
      setSubmitError('Vui lòng chọn loại dự án: Web App, API hoặc Web + API.')
      return
    }
    try {
      const project = createProject({
        name: form.name,
        description: form.description,
        type: form.type,
        webFolders: [],
        apiFolders: [],
      })
      // Navigation is intentionally local/immediate. Persistence continues in the background.
      navigate(`/project/${project.id}`, { replace: false })
    } catch (error) {
      setSubmitError(error?.message || 'Không thể tạo dự án. Vui lòng thử lại.')
    }
  }

  return (
    <AppShell>
      <Breadcrumb items={['Trang chủ', 'Bảng điều khiển']} />
      <PageHeader index="01" title="Bảng điều khiển" subtitle="Tạo và quản lý dự án kiểm thử." />

      <div className="dashboard-grid">
        <section className="card create-project-card" id="new-project-form">
          <div className="section-title">
            <span className="section-icon">1</span>
            <div><h2>Tạo dự án mới</h2><p>Các folder màn hình sẽ được tạo tự động sau khi phân tích tài liệu.</p></div>
          </div>

          <form onSubmit={submit} noValidate>
            <label className="field-label">Tên dự án <em>*</em></label>
            <input
              ref={nameInputRef}
              className="input"
              value={form.name}
              onChange={e => setForm({ ...form, name: e.target.value })}
              placeholder="Ví dụ: Trái phiếu BIDV"
            />

            <label className="field-label">Mô tả ngắn</label>
            <textarea
              className="textarea"
              maxLength={500}
              value={form.description}
              onChange={e => setForm({ ...form, description: e.target.value })}
              placeholder="Mô tả dự án..."
            />
            <div className="counter">{form.description.length}/500</div>

            <label className="field-label">Loại dự án <em>*</em></label>
            <select className="input" value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}>
              <option value="">Chọn loại dự án</option>
              <option value="web">Web App</option>
              <option value="api">API</option>
              <option value="both">Web + API</option>
            </select>

            {submitError && <div className="compact-warning danger">{submitError}</div>}

            <div className="actions-end">
              <button type="submit" className="btn btn-primary btn-balanced">Tạo dự án <ArrowRight size={18} /></button>
            </div>
          </form>
        </section>

        <aside className="right-stack">
          <section className="card recent-card">
            <div className="recent-head">
              <h3><FolderOpen size={21} /> Dự án gần đây</h3>
            </div>
            {projects.slice(0, 5).map(project => (
              <button key={project.id} className="recent-item" onClick={() => navigate(`/project/${project.id}`)}>
                <span className="folder-swatch"><FolderOpen size={20} /></span>
                <span><b>{project.name}</b><small>Cập nhật {project.updated || 'vừa xong'}</small></span>
                <i>•••</i>
              </button>
            ))}
          </section>
        </aside>
      </div>
    </AppShell>
  )
}
