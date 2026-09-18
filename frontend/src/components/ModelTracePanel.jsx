import { useCallback, useEffect, useState } from 'react'
import { Bug, ChevronDown, ChevronUp, Clipboard, Trash2, X } from 'lucide-react'
import { MODEL_TRACE_EVENT, clearModelTrace, getModelTrace } from '../api/client'

const formatTime = value => value ? new Date(value).toLocaleTimeString('vi-VN') : '—'

export default function ModelTracePanel() {
  const [open, setOpen] = useState(false)
  const [trace, setTrace] = useState(() => getModelTrace())

  const refresh = useCallback(() => setTrace(getModelTrace()), [])

  useEffect(() => {
    window.addEventListener(MODEL_TRACE_EVENT, refresh)
    return () => window.removeEventListener(MODEL_TRACE_EVENT, refresh)
  }, [refresh])

  const copy = async value => {
    if (!value) return
    try {
      await navigator.clipboard.writeText(value)
    } catch {
      // Clipboard access can be blocked in an embedded browser.
    }
  }

  const stateLabel = trace?.status === 'streaming' ? 'Đang nhận stream' : trace?.status === 'completed' ? 'Hoàn tất' : trace?.status === 'failed' ? 'Lỗi' : 'Chưa có request'

  return (
    <div className="model-trace-root">
      {open && (
        <aside className="model-trace-panel" aria-label="Theo dõi request model">
          <div className="model-trace-header">
            <div>
              <div className="model-trace-title"><Bug size={16} /> Model trace</div>
              <div className={`model-trace-status ${trace?.status || 'idle'}`}>{stateLabel}</div>
            </div>
            <div className="model-trace-actions">
              <button className="model-trace-icon-button" type="button" title="Xóa trace" onClick={() => { clearModelTrace(); refresh() }}><Trash2 size={15} /></button>
              <button className="model-trace-icon-button" type="button" title="Đóng" onClick={() => setOpen(false)}><X size={16} /></button>
            </div>
          </div>

          {trace ? (
            <>
              <div className="model-trace-meta">
                <span>{trace.title}</span><span>{trace.model || 'Chưa chọn model'}</span>
                <span>{formatTime(trace.startedAt)} → {formatTime(trace.finishedAt)}</span>
              </div>
              {trace.baseUrl && <div className="model-trace-url" title={trace.baseUrl}>{trace.baseUrl}</div>}
              {trace.error && <div className="model-trace-error">{trace.error}</div>}
              <TraceSection label={`Input${trace.inputTruncated ? ' (đã rút gọn)' : ''}`} value={trace.input} onCopy={copy} />
              <TraceSection label={`Output${trace.outputTruncated ? ' (đã rút gọn)' : ''}`} value={trace.output} onCopy={copy} empty={trace.status === 'streaming' ? 'Đang chờ model trả nội dung…' : 'Chưa có output.'} />
            </>
          ) : <div className="model-trace-empty">Chưa có lần gọi model nào trong session này.</div>}
        </aside>
      )}
      <button className="model-trace-toggle" type="button" onClick={() => setOpen(value => !value)} aria-expanded={open}>
        <Bug size={16} /> Model trace {open ? <ChevronDown size={15} /> : <ChevronUp size={15} />}
      </button>
    </div>
  )
}

function TraceSection({ label, value, onCopy, empty = 'Chưa có dữ liệu.' }) {
  return (
    <section className="model-trace-section">
      <div className="model-trace-section-head">
        <span>{label}</span>
        <button type="button" onClick={() => onCopy(value)} disabled={!value} title="Sao chép"><Clipboard size={14} /> Sao chép</button>
      </div>
      <pre>{value || empty}</pre>
    </section>
  )
}
