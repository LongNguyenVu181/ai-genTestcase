import { ArrowLeft, ArrowRight, File, FileJson2, FileText, Trash2, UploadCloud } from 'lucide-react'
import { useEffect, useRef, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import { analyzeApi, analyzeWeb, getAnalysisJob, invalidateProjectCache, loadAiConfig } from '../api/client'
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
  const [jobStatus, setJobStatus] = useState(null)
  const pollTokenRef = useRef(0)

  const config = loadAiConfig()
  const hasFiles = scope === 'api' ? !!designFile && !!baFile : files.length > 0
  const addFiles = fileList => setFiles(prev => [...prev, ...Array.from(fileList)])
  const jobStorageKey = `testpilot.job.${projectId}.${folderId}.${scope}`
  const runStorageKey = `testpilot.run.${projectId}.${folderId}`

  const finishJob = job => {
    sessionStorage.removeItem(jobStorageKey)
    invalidateProjectCache(projectId)
    if (job?.run_id) sessionStorage.setItem(runStorageKey, job.run_id)
    setJobStatus(job)
    setPhase('done')
    if (params.folderId) {
      const runQuery = job?.run_id ? `&run=${encodeURIComponent(job.run_id)}` : ''
      navigate(`/project/${projectId}/folder/${folderId}/testcases?scope=${encodeURIComponent(scope)}${runQuery}`, { replace: true })
    } else {
      navigate(`/project/${projectId}/scope/${scope}`, { replace: true })
    }
  }

  const pollJob = async jobId => {
    const token = ++pollTokenRef.current
    let transientFailures = 0
    setPhase('analyzing')
    setJobStatus(current => current || { status: 'queued', stage: 'queued', progress: 0, message: 'Đã gửi tài liệu. Đang khởi tạo tác vụ...' })

    while (token === pollTokenRef.current) {
      try {
        const job = await getAnalysisJob(jobId)
        transientFailures = 0
        setJobStatus(job)
        if (job.status === 'completed') {
          finishJob(job)
          return
        }
        if (job.status === 'failed') {
          sessionStorage.removeItem(jobStorageKey)
          setError(job.error?.message || job.message || 'Phân tích tài liệu thất bại.')
          setPhase('idle')
          return
        }
      } catch (e) {
        transientFailures += 1
        if (e.status >= 500 && transientFailures <= 10) {
          setJobStatus(current => ({
            ...(current || {}),
            message: 'Kết nối tạm thời gián đoạn. Đang tự kết nối lại...',
          }))
        } else {
          setError(e.message || 'Không kiểm tra được trạng thái phân tích.')
          setPhase('idle')
          return
        }
      }
      await new Promise(resolve => window.setTimeout(resolve, 3500))
    }
  }

  useEffect(() => {
    const existingJob = sessionStorage.getItem(jobStorageKey)
    if (existingJob) pollJob(existingJob)
    return () => { pollTokenRef.current += 1 }
    // Resume exactly the job for this project/scope after refresh.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [jobStorageKey])

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
    setJobStatus({ status: 'queued', stage: 'uploading', progress: 0, message: 'Đang tải tài liệu lên server...' })
    try {
      const result = scope === 'api'
        ? await analyzeApi({ designFile, baFile, config, projectId, folderId })
        : await analyzeWeb({ files, config, projectId, folderId })

      if (!result.job_id) throw new Error('Server không trả về mã tác vụ phân tích.')
      sessionStorage.setItem(jobStorageKey, result.job_id)
      setJobStatus({ status: result.status || 'queued', stage: 'queued', progress: 0, message: result.message || 'Đã tiếp nhận tài liệu.' })
      await pollJob(result.job_id)
    } catch (e) {
      setError(e.message || 'Không thể khởi tạo tác vụ phân tích.')
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
          <div>
            <div className="api-target-note">
              <FileJson2 size={20} />
              <div>
                <b>API Design là phạm vi mục tiêu</b>
                <span>AI sẽ nhận diện endpoint từ API Design, sau đó chỉ chọn các phần liên quan trong tài liệu BA để phân tích sâu. Các API khác trong BA sẽ không tự sinh testcase.</span>
              </div>
            </div>
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
          </div>
        )}
      </section>

      {!!displayFiles.length && (
        <section className="card file-table-card">
          <div className="file-table-head">
            <div><File size={22} /><span><h3>Tài liệu đã chọn ({displayFiles.length})</h3></span></div>
            <button className="btn btn-outline btn-balanced" onClick={() => { setFiles([]); setDesignFile(null); setBaFile(null); setPhase('idle'); setJobStatus(null); setError('') }}>
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
        <div className="analysis-toast analysis-toast-job">
          <span className="spinner"></span>
          <div className="analysis-toast-content">
            <b>{jobStatus?.message || 'Đang phân tích tài liệu...'}</b>
            <small>{Math.max(0, Number(jobStatus?.progress || 0))}% · AI chạy nền, bạn có thể giữ hoặc tải lại trang.</small>
            <div className="analysis-progress"><span style={{ width: `${Math.max(3, Number(jobStatus?.progress || 0))}%` }} /></div>
          </div>
        </div>
      )}
    </AppShell>
  )
}
