import { ArrowLeft, ArrowRight, File, FileJson2, FileText, Trash2, UploadCloud } from 'lucide-react'
import { useRef, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import { analyzeApi, analyzeWeb, loadAiConfig } from '../api/client'
import { getProjectById } from '../data/projectStore'

const fileMeta = file => ({
  name: file.name,
  type: file.name.split('.').pop()?.toUpperCase() || 'FILE',
  size: file.size > 1024 * 1024 ? `${(file.size / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(file.size / 1024))} KB`,
  status: 'Sẵn sàng',
})

export default function UploadPage() {
  const navigate = useNavigate()
  const params = useParams()
  const { projectId = 'website-tmdt' } = params
  const [searchParams] = useSearchParams()
  const scope = params.scope === 'api' || searchParams.get('scope') === 'api' ? 'api' : 'web'
  const folderId = params.folderId || `__scope_${scope}__`
  const project = getProjectById(projectId)
  const inputRef = useRef(null)
  const designRef = useRef(null)
  const baRef = useRef(null)

  const [files, setFiles] = useState([])
  const [designFile, setDesignFile] = useState(null)
  const [baFile, setBaFile] = useState(null)
  const [phase, setPhase] = useState('idle')
  const [error, setError] = useState('')

  const config = loadAiConfig()
  const hasFiles = scope === 'api' ? !!designFile && !!baFile : files.length > 0
  const addFiles = fileList => setFiles(prev => [...prev, ...Array.from(fileList)])

  const ensureConfig = () => {
    if (!config.apiKey?.trim()) {
      setError('Chưa có API Key. Hãy cấu hình API trước khi phân tích.')
      return false
    }
    return true
  }

  const analyze = async () => {
    if (!hasFiles || !ensureConfig()) return
    setError('')
    setPhase('analyzing')
    try {
      const result = scope === 'api'
        ? await analyzeApi({ designFile, baFile, config, projectId, folderId })
        : await analyzeWeb({ files, config, projectId, folderId })

      sessionStorage.setItem(`testpilot.run.${projectId}.${folderId}`, result.run_id)
      navigate(`/project/${projectId}/scope/${scope}`, { replace: true })
    } catch (e) {
      setError(e.message || 'Phân tích tài liệu thất bại.')
      setPhase('idle')
    }
  }

  const displayFiles = (scope === 'api' ? [designFile, baFile].filter(Boolean) : files).map(fileMeta)

  return (
    <AppShell wide>
      <div className="clean-page-head upload-clean-head">
        <button className="icon-text-button" onClick={() => navigate(`/project/${projectId}`)}>
          <ArrowLeft size={18} /> Quay lại
        </button>
        <div className="clean-title">
          <span className="clean-title-icon"><UploadCloud size={27} /></span>
          <div><h1>Tải tài liệu</h1><p>{project?.name || projectId} · {scope === 'api' ? 'API' : 'Web App'}</p></div>
        </div>
        <div />
      </div>

      {error && <div className="pipeline-alert danger"><b>Không thể tiếp tục</b><span>{error}</span></div>}
      {!config.apiKey && (
        <div className="pipeline-alert warning">
          <b>Chưa cấu hình API</b>
          <button className="btn btn-outline btn-balanced" onClick={() => navigate('/api-keys')}>Mở cấu hình</button>
        </div>
      )}

      <section className="card upload-card upload-card-clean">
        {scope === 'web' ? (
          <div className="dropzone clean-dropzone" onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); addFiles(e.dataTransfer.files) }}>
            <FileText size={48} />
            <h3>Kéo thả tài liệu vào đây</h3>
            <p>PDF, MD, TXT, DOC, DOCX</p>
            <button className="btn btn-primary btn-balanced" onClick={() => inputRef.current?.click()}>Chọn tài liệu</button>
            <input ref={inputRef} type="file" multiple accept=".pdf,.md,.txt,.doc,.docx" hidden onChange={e => addFiles(e.target.files)} />
          </div>
        ) : (
          <div className="api-upload-grid">
            <button className={`api-upload-slot ${designFile ? 'has-file' : ''}`} onClick={() => designRef.current?.click()}>
              <FileJson2 size={36} /><b>API Design / Spec</b><span>{designFile ? designFile.name : 'Chọn tài liệu'}</span>
            </button>
            <button className={`api-upload-slot ${baFile ? 'has-file' : ''}`} onClick={() => baRef.current?.click()}>
              <FileText size={36} /><b>BA / Business</b><span>{baFile ? baFile.name : 'Chọn tài liệu'}</span>
            </button>
            <input ref={designRef} type="file" accept=".json,.yaml,.yml,.pdf,.md,.txt,.doc,.docx" hidden onChange={e => setDesignFile(e.target.files?.[0] || null)} />
            <input ref={baRef} type="file" accept=".pdf,.md,.txt,.doc,.docx" hidden onChange={e => setBaFile(e.target.files?.[0] || null)} />
          </div>
        )}
      </section>

      {!!displayFiles.length && (
        <section className="card file-table-card">
          <div className="file-table-head">
            <div><File size={22} /><span><h3>Tài liệu đã chọn ({displayFiles.length})</h3></span></div>
            <button className="btn btn-outline btn-balanced" onClick={() => { setFiles([]); setDesignFile(null); setBaFile(null); setPhase('idle') }}>
              <Trash2 size={17} /> Xóa tất cả
            </button>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Tên tài liệu</th><th>Loại</th><th>Dung lượng</th><th>Trạng thái</th></tr></thead>
              <tbody>
                {displayFiles.map((file, idx) => (
                  <tr key={`${file.name}-${idx}`}>
                    <td><b>{file.name}</b></td><td>{file.type}</td><td>{file.size}</td><td><span className="status neutral">● {file.status}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <div className="actions-end page-actions clean-actions">
        <button className="btn btn-outline btn-balanced" onClick={() => navigate(-1)}><ArrowLeft size={18} /> Quay lại</button>
        <button className="btn btn-primary btn-balanced" disabled={phase === 'analyzing' || !hasFiles} onClick={analyze}>
          {phase === 'analyzing' ? 'Đang phân tích...' : 'Phân tích tài liệu'}
          {phase !== 'analyzing' && <ArrowRight size={18} />}
        </button>
      </div>

      {phase === 'analyzing' && (
        <div className="analysis-toast">
          <span className="spinner"></span>
          <div><b>Đang phân tích tài liệu...</b><small>Vui lòng giữ trang này mở.</small></div>
        </div>
      )}
    </AppShell>
  )
}
