import { ArrowLeft, ChevronDown, ChevronUp, Download, Edit3, Filter, Plus, Search, ShieldCheck, Trash2 } from 'lucide-react'
import { useEffect, useMemo, useState } from 'react'
import { useNavigate, useParams, useSearchParams } from 'react-router-dom'
import AppShell from '../components/AppShell'
import Modal from '../components/Modal'
import { getProjectById } from '../data/projectStore'
import {
  createTestcase,
  deleteTestcase,
  folderDownloadUrl,
  getFolderTestcases,
  getRun,
  getScreenTestcases,
  screenDownloadUrl,
  updateTestcase,
} from '../api/client'

const TYPES = ['Giao diện', 'Kiểm tra dữ liệu', 'Chức năng', 'Ngoại lệ', 'Popup', 'Luồng']
const slug = value => String(value || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g,'d').replace(/[^a-z0-9]+/g, '-')
const CATEGORY_ORDER = { VALIDATION: 1, REQUEST_VALIDATION: 1, RESPONSE_VALIDATION: 1, UI: 2, ACTION: 3, DATA_GRID: 4, BUSINESS_FLOW: 5, EXCEPTION: 6 }
const FEATURE_GROUP_ORDER = { PRECONDITION_PERMISSION: 1, GENERAL_UI: 2, FILTER: 3, DATA_GRID: 4, FUNCTION: 5, API: 6 }

const blankAdvanced = {
  actualResult: '', run1: '', run2: '', run3: '', currentResult: '', note: '', errorCode: '',
  qcWriter: '', sprintWriter: '', qcExecutor: '', sprintExecutor: '', reviewer: '', reviewDate: '',
  reviewContent: '', needAuto: '', automated: '', smoke: '', regression: '', outdated: '', outdatedDate: '',
}

const makeDraft = (items, item) => item
  ? { ...item, stepsText: (item.steps || []).join('\n') }
  : {
      id: `TC_${String(items.length + 1).padStart(3, '0')}`,
      name: '',
      type: 'Chức năng',
      preCondition: '',
      importance: '',
      testData: '',
      stepsText: '',
      expectedResult: '',
      sourceRuleId: '',
      sourceRequirement: '',
      featureGroup: 'FUNCTION',
      featureName: 'Testcase thủ công',
      category: 'ACTION',
      screen: '',
      ...blankAdvanced,
    }

const downloadBlob = (name, type, content) => {
  const blob = new Blob([content], { type })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  URL.revokeObjectURL(url)
}

export default function TestcasePage() {
  const navigate = useNavigate()
  const params = useParams()
  const { projectId = 'website-tmdt' } = params
  const [searchParams] = useSearchParams()
  const scope = params.scope === 'api' || searchParams.get('scope') === 'api' ? 'api' : 'web'
  const folderId = params.folderId || `__scope_${scope}__`
  const screenFilter = searchParams.get('screen') || ''
  const project = getProjectById(projectId)
  const runFromUrl = searchParams.get('run')
  const savedRun = sessionStorage.getItem(`testpilot.run.${projectId}.${folderId}`)
  const runId = runFromUrl || savedRun

  const [items, setItems] = useState([])
  const [run, setRun] = useState(null)
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState('')
  const [openId, setOpenId] = useState(null)
  const [query, setQuery] = useState('')
  const [filter, setFilter] = useState('')
  const [editorOpen, setEditorOpen] = useState(false)
  const [deleteItem, setDeleteItem] = useState(null)
  const [draft, setDraft] = useState(null)
  const [saving, setSaving] = useState(false)

  const reload = async () => {
    setLoading(true)
    setLoadError('')
    try {
      let result
      let loadedRun = null

      if (screenFilter) {
        result = await getScreenTestcases({ projectId, scope, screen: screenFilter })
      } else if (runId) {
        try {
          result = await getRun(runId)
          loadedRun = result
        } catch {
          result = await getFolderTestcases({ projectId, folderId, scope })
        }
      } else {
        result = await getFolderTestcases({ projectId, folderId, scope })
      }

      const next = result.testcases || []
      if (next.length === 0) {
        navigate(`/project/${projectId}/scope/${scope}`, { replace: true })
        return
      }

      setRun(loadedRun)
      setItems(next)
      setOpenId(prev => prev || next[0]?.recordId || null)
    } catch (e) {
      setLoadError(e.message || 'Không tải được testcase từ backend.')
      setItems([])
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { reload() }, [runId, projectId, folderId, scope, screenFilter])

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase()
    return items
      .filter(item => (!screenFilter || item.screen === screenFilter) && (!filter || item.type === filter) && (!q || item.id.toLowerCase().includes(q) || item.name.toLowerCase().includes(q)))
      .sort((a, b) => {
        const fg = (FEATURE_GROUP_ORDER[a.featureGroup] ?? 99) - (FEATURE_GROUP_ORDER[b.featureGroup] ?? 99)
        if (fg) return fg
        const fn = String(a.featureName || '').localeCompare(String(b.featureName || ''), 'vi')
        if (fn) return fn
        const cat = (CATEGORY_ORDER[a.category] ?? 99) - (CATEGORY_ORDER[b.category] ?? 99)
        if (cat) return cat
        return (a.sortOrder ?? 0) - (b.sortOrder ?? 0)
      })
  }, [items, query, filter, screenFilter])

  const groups = useMemo(() => {
    const result = []
    const map = new Map()
    for (const item of visible) {
      const key = `${item.featureGroup || 'OTHER'}::${item.featureName || 'Khác'}`
      if (!map.has(key)) {
        const group = {
          key,
          featureGroup: item.featureGroup || 'OTHER',
          featureName: item.featureName || 'Khác',
          items: [],
        }
        map.set(key, group)
        result.push(group)
      }
      map.get(key).items.push(item)
    }
    return result
  }, [visible])

  const openEditor = item => {
    setDraft(makeDraft(items, item))
    setEditorOpen(true)
  }

  const normalizeDraft = draftValue => {
    const { stepsText, recordId, runId: _runId, projectId: _projectId, folderId: _folderId, scope: _scope, validation, sortOrder, ...rest } = draftValue
    return {
      ...blankAdvanced,
      ...rest,
      steps: String(stepsText || '').split('\n').map(x => x.trim()).filter(Boolean),
      screen: rest.screen || screenFilter || project?.name || projectId,
    }
  }

  const save = async () => {
    if (!draft) return
    setSaving(true)
    try {
      const testcase = normalizeDraft(draft)
      const result = draft.recordId
        ? await updateTestcase({ recordId: draft.recordId, testcase })
        : await createTestcase({ projectId, folderId, scope, runId, testcase })

      setOpenId(result.testcase.recordId)
      setEditorOpen(false)
      await reload()
    } catch (e) {
      alert(e.message || 'Không lưu được testcase.')
    } finally {
      setSaving(false)
    }
  }

  const confirmDelete = async () => {
    if (!deleteItem?.recordId) return
    try {
      await deleteTestcase(deleteItem.recordId)
      setDeleteItem(null)
      await reload()
    } catch (e) {
      alert(e.message || 'Không xóa được testcase.')
    }
  }

  const excelUrl = screenFilter
    ? screenDownloadUrl({ projectId, scope, screen: screenFilter, type: 'excel' })
    : folderDownloadUrl({ projectId, folderId, scope, type: 'excel' })
  const xmindUrl = screenFilter
    ? screenDownloadUrl({ projectId, scope, screen: screenFilter, type: 'xmind' })
    : folderDownloadUrl({ projectId, folderId, scope, type: 'xmind' })

  return (
    <AppShell wide>
      <div className="testcase-workspace-page">
        <div className="clean-page-head testcase-clean-head">
          <button className="icon-text-button" onClick={() => navigate(`/project/${projectId}/scope/${scope}`)}>
            <ArrowLeft size={18} /> Quay lại
          </button>
          <div className="clean-title">
            <div>
              <h1>{screenFilter || 'Testcase'}</h1>
              <p>{visible.length} testcase</p>
            </div>
          </div>
          <div />
        </div>

        {loadError && <div className="pipeline-alert danger"><b>Không đọc được testcase</b><span>{loadError}</span></div>}

        <section className="test-toolbar">
          <label className="search-box large"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Tìm kiếm theo ID, tên testcase..." /></label>
          <label className="select-icon"><Filter size={17} /><select value={filter} onChange={e => setFilter(e.target.value)}><option value="">Tất cả loại testcase</option>{TYPES.map(t => <option key={t}>{t}</option>)}</select></label>
          <div className="test-toolbar-actions">
            <button className="btn btn-primary" onClick={() => openEditor(null)}><Plus size={18} /> Thêm testcase</button>
            <a className="btn btn-outline" href={excelUrl}><Download size={17} /> Tải Excel</a>
            <a className="btn btn-outline" href={xmindUrl}><Download size={17} /> Tải XMind</a>
          </div>
        </section>

        {loading ? (
          <div className="workspace-loading"><span className="spinner"></span><b>Đang tải testcase...</b></div>
        ) : !items.length ? (
          <div className="workspace-empty">
            <ShieldCheck size={38} />
            <h3>Chưa có testcase</h3>
          </div>
        ) : (
          <div className="tc-review-groups">
            {groups.map(group => (
              <section className="tc-review-group" key={group.key}>
                <div className="tc-review-group-head">
                  <div>
                    <small>{group.featureGroup}</small>
                    <h3>{group.featureName}</h3>
                  </div>
                  <span>{group.items.length} testcase</span>
                </div>

                <div className="accordion-list">
                  {group.items.map(item => {
                    const open = openId === item.recordId
                    const valid = item.validation?.valid !== false
                    return (
                      <article className={`testcase-card ${open ? 'open' : ''} ${valid ? '' : 'invalid'}`} key={item.recordId}>
                        <button className="testcase-head" onClick={() => setOpenId(open ? null : item.recordId)}>
                          <span className="tc-id">{item.id}</span>
                          <b>{item.name}</b>
                          <span className={`type-chip type-${slug(item.type)}`}>{item.type}</span>
                          <span className={`validation-dot ${valid ? 'ok' : 'bad'}`} title={valid ? 'Hợp lệ' : 'Thiếu dữ liệu'}></span>
                          {open ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                        </button>

                        {open && <div className="testcase-body excel-review-order">
                          {!valid && <div className="field-validation-box">
                            <b>Validation testcase</b>
                            {Object.entries(item.validation?.errors || {}).map(([field, message]) => <span key={field}>• {field}: {message}</span>)}
                          </div>}

                          <div className="detail-block"><h4>PreConditions</h4><div className="detail-box preline">{item.preCondition || '—'}</div></div>
                          <div className="detail-block"><h4>Importance</h4><div className="detail-box">{item.importance || '—'}</div></div>
                          <div className="detail-block detail-span-2"><h4>Step</h4><div className="detail-box steps">{(item.steps || []).map((step, i) => <div key={i}><span>{i+1}.</span> {step}</div>)}</div></div>
                          <div className="detail-block"><h4>Data test</h4><div className="detail-box preline">{item.testData || '—'}</div></div>
                          <div className="detail-block"><h4>Expected Result</h4><div className="detail-box preline">{item.expectedResult || '—'}</div></div>

                          <details className="execution-review detail-span-2">
                            <summary>Thông tin thực thi / review / automation</summary>
                            <div className="execution-review-grid">
                              <span><b>Actual Result</b>{item.actualResult || '—'}</span>
                              <span><b>Lần 1 / 2 / 3</b>{[item.run1,item.run2,item.run3].filter(Boolean).join(' / ') || '—'}</span>
                              <span><b>Kết quả hiện tại</b>{item.currentResult || '—'}</span>
                              <span><b>Ghi chú</b>{item.note || '—'}</span>
                              <span><b>Mã lỗi</b>{item.errorCode || '—'}</span>
                              <span><b>QC viết testcase</b>{item.qcWriter || '—'}</span>
                              <span><b>Sprint viết testcase</b>{item.sprintWriter || '—'}</span>
                              <span><b>QC thực hiện test</b>{item.qcExecutor || '—'}</span>
                              <span><b>Sprint thực hiện test</b>{item.sprintExecutor || '—'}</span>
                              <span><b>Người review</b>{item.reviewer || '—'}</span>
                              <span><b>Ngày review</b>{item.reviewDate || '—'}</span>
                              <span><b>Nội dung review</b>{item.reviewContent || '—'}</span>
                            </div>
                          </details>

                          <div className="testcase-actions detail-span-2">
                            <button className="btn btn-success" onClick={() => openEditor(item)}><Edit3 size={17} /> Sửa</button>
                            <button className="btn btn-danger" onClick={() => setDeleteItem(item)}><Trash2 size={17} /> Xóa</button>
                          </div>
                        </div>}
                      </article>
                    )
                  })}
                </div>
              </section>
            ))}
          </div>
        )}

        <Modal open={editorOpen} title={draft?.recordId ? 'Sửa testcase' : 'Thêm testcase'} onClose={() => setEditorOpen(false)} width="1040px">
          {draft && <div className="modal-body testcase-form">
            <div className="form-grid two">
              <label><span>ID <small>(tự động TC_001 → TC_N)</small></span><input className="input" value={draft.id} readOnly /></label>
              <label><span>Loại testcase</span><select className="input" value={draft.type} onChange={e => setDraft({...draft,type:e.target.value})}>{TYPES.map(t => <option key={t}>{t}</option>)}</select></label>
            </div>
            <label><span>Name *</span><input className="input" value={draft.name} onChange={e => setDraft({...draft,name:e.target.value})} /></label>
            <div className="form-grid two">
              <label><span>PreConditions</span><textarea className="textarea small-area" value={draft.preCondition} onChange={e => setDraft({...draft,preCondition:e.target.value})} /></label>
              <label><span>Importance</span><input className="input" value={draft.importance} onChange={e => setDraft({...draft,importance:e.target.value})} placeholder="Để trống nếu chưa có quy định" /></label>
            </div>
            <label><span>Step * <small>(mỗi dòng là một bước)</small></span><textarea className="textarea" value={draft.stepsText} onChange={e => setDraft({...draft,stepsText:e.target.value})} /></label>
            <div className="form-grid two">
              <label><span>Data test</span><textarea className="textarea small-area" value={draft.testData} onChange={e => setDraft({...draft,testData:e.target.value})} /></label>
              <label><span>Expected Result *</span><textarea className="textarea small-area" value={draft.expectedResult} onChange={e => setDraft({...draft,expectedResult:e.target.value})} /></label>
            </div>

            <details className="advanced-edit">
              <summary>Phần còn lại theo Excel: thực thi / review / automation</summary>
              <div className="form-grid two">
                <label><span>Actual Result</span><textarea className="textarea small-area" value={draft.actualResult} onChange={e => setDraft({...draft,actualResult:e.target.value})} /></label>
                <label><span>Kết quả hiện tại</span><input className="input" value={draft.currentResult} onChange={e => setDraft({...draft,currentResult:e.target.value})} /></label>
                <label><span>Lần 1</span><input className="input" value={draft.run1} onChange={e => setDraft({...draft,run1:e.target.value})} /></label>
                <label><span>Lần 2</span><input className="input" value={draft.run2} onChange={e => setDraft({...draft,run2:e.target.value})} /></label>
                <label><span>Lần 3</span><input className="input" value={draft.run3} onChange={e => setDraft({...draft,run3:e.target.value})} /></label>
                <label><span>Ghi chú</span><input className="input" value={draft.note} onChange={e => setDraft({...draft,note:e.target.value})} /></label>
                <label><span>Mã lỗi</span><input className="input" value={draft.errorCode} onChange={e => setDraft({...draft,errorCode:e.target.value})} /></label>
                <label><span>QC viết testcase</span><input className="input" value={draft.qcWriter} onChange={e => setDraft({...draft,qcWriter:e.target.value})} /></label>
                <label><span>Sprint viết testcase</span><input className="input" value={draft.sprintWriter} onChange={e => setDraft({...draft,sprintWriter:e.target.value})} /></label>
                <label><span>QC thực hiện test</span><input className="input" value={draft.qcExecutor} onChange={e => setDraft({...draft,qcExecutor:e.target.value})} /></label>
                <label><span>Sprint thực hiện test</span><input className="input" value={draft.sprintExecutor} onChange={e => setDraft({...draft,sprintExecutor:e.target.value})} /></label>
                <label><span>Người review</span><input className="input" value={draft.reviewer} onChange={e => setDraft({...draft,reviewer:e.target.value})} /></label>
                <label><span>Ngày review</span><input className="input" value={draft.reviewDate} onChange={e => setDraft({...draft,reviewDate:e.target.value})} /></label>
                <label className="form-span-2"><span>Nội dung review</span><textarea className="textarea small-area" value={draft.reviewContent} onChange={e => setDraft({...draft,reviewContent:e.target.value})} /></label>
              </div>
            </details>

            <div className="actions-end">
              <button className="btn btn-outline" onClick={() => setEditorOpen(false)}>Hủy</button>
              <button className="btn btn-primary" disabled={saving} onClick={save}>{saving ? 'Đang lưu...' : 'Lưu vào DB'}</button>
            </div>
          </div>}
        </Modal>
      </div>

      <Modal open={!!deleteItem} title="Xác nhận xóa testcase" onClose={() => setDeleteItem(null)}>
        <div className="modal-body"><p>Bạn có chắc muốn xóa testcase <b>{deleteItem?.id}</b>? Dữ liệu sẽ bị xóa khỏi DB.</p><div className="actions-end"><button className="btn btn-outline" onClick={() => setDeleteItem(null)}>Hủy</button><button className="btn btn-danger" onClick={confirmDelete}>Xóa</button></div></div>
      </Modal>
    </AppShell>
  )
}
