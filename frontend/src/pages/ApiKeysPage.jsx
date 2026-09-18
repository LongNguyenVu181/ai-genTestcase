import { CheckCircle2, KeyRound, PlugZap, RotateCcw, Save, ShieldCheck } from 'lucide-react'
import { useMemo, useState } from 'react'
import AppShell from '../components/AppShell'
import Breadcrumb from '../components/Breadcrumb'
import PageHeader from '../components/PageHeader'
import { DEFAULT_CONFIG, loadAiConfig, maskKey, saveAiConfig, testAiConfig } from '../api/client'

const models = [
  { value: 'qwen-max', label: 'Qwen Max', note: 'Model mặc định đang dùng trong pipeline hiện tại.' },
  { value: 'qwen3.8-max', label: 'Qwen 3.8 Max', note: 'Model Max phiên bản mới để thử prompt.' },
  { value: 'qwen-plus', label: 'Qwen Plus', note: 'Nhẹ hơn, phù hợp test nhanh.' },
]

export default function ApiKeysPage() {
  const [config, setConfig] = useState(() => loadAiConfig())
  const [status, setStatus] = useState({ type: '', text: '' })
  const [testing, setTesting] = useState(false)

  const selected = useMemo(() => models.find(x => x.value === config.model) || models[0], [config.model])

  const save = () => {
    saveAiConfig(config)
    setStatus({ type: 'success', text: 'Đã lưu cấu hình cho phiên làm việc hiện tại.' })
  }

  const resetDefault = () => {
    const next = { ...DEFAULT_CONFIG }
    setConfig(next)
    saveAiConfig(next)
    setStatus({ type: 'success', text: 'Đã khôi phục endpoint/model mặc định của pipeline.' })
  }

  const test = async () => {
    if (!config.apiKey.trim()) {
      setStatus({ type: 'danger', text: 'Vui lòng nhập API Key trước.' })
      return
    }
    setTesting(true)
    setStatus({ type: '', text: '' })
    try {
      const result = await testAiConfig(config)
      saveAiConfig(config)
      setStatus({ type: 'success', text: `Kết nối thành công ${result.model} (${result.elapsed?.toFixed?.(2) ?? result.elapsed}s).` })
    } catch (e) {
      setStatus({ type: 'danger', text: e.message || 'Kết nối thất bại.' })
    } finally {
      setTesting(false)
    }
  }

  return (
    <AppShell>
      <Breadcrumb items={['Trang chủ', 'Quản lý khóa API']} />
      <PageHeader index="05" title="Quản lý khóa API" subtitle="Cấu hình Qwen model và DashScope endpoint dùng cho pipeline phân tích testcase." />

      <section className="card model-card">
        <div className="section-title compact"><KeyRound size={25} /><div><h2>Mô hình AI mặc định</h2><p>Model này sẽ được dùng cho pipeline phân tích tài liệu và sinh testcase.</p></div></div>
        <div className="model-grid">
          {models.map(item => (
            <button key={item.value} className={`model-option ${config.model === item.value ? 'active' : ''}`} onClick={() => setConfig({ ...config, model: item.value })}>
              <span className="model-logo">Q</span><b>{item.label}</b><small>{item.note}</small><i>{config.model === item.value ? '●' : '○'}</i>
            </button>
          ))}
        </div>
      </section>

      <section className="card api-table-card config-form-card">
        <div className="section-title compact"><ShieldCheck size={25} /><div><h2>Cấu hình kết nối Qwen</h2><p>API key không ghi vào source code hoặc database. Key chỉ dùng trong phiên hiện tại và bị xóa khi refresh trang.</p></div></div>
        <div className="config-grid">
          <label><span className="field-label">Nhà cung cấp</span><input className="input" value="Qwen / DashScope" disabled /></label>
          <label><span className="field-label">Model</span><input className="input" value={selected.label} disabled /></label>
          <label className="config-full"><span className="field-label">Base URL Endpoint</span><input className="input" value={config.baseUrl} onChange={e => setConfig({ ...config, baseUrl: e.target.value })} placeholder={DEFAULT_CONFIG.baseUrl} /></label>
          <label className="config-full"><span className="field-label">API Key</span><input className="input" type="password" value={config.apiKey} onChange={e => setConfig({ ...config, apiKey: e.target.value })} placeholder="Paste API key dùng để test prompt" /></label>
        </div>
        <div className="api-config-summary"><span><b>Model:</b> {config.model}</span><span><b>Key:</b> {maskKey(config.apiKey)}</span></div>
        {status.text && <div className={`config-status ${status.type}`}><CheckCircle2 size={17} /> {status.text}</div>}
        <div className="actions-end">
          <button className="btn btn-outline" disabled={testing} onClick={resetDefault}><RotateCcw size={17} /> Khôi phục mặc định</button>
          <button className="btn btn-outline" disabled={testing} onClick={test}><PlugZap size={17} /> {testing ? 'Đang kiểm tra...' : 'Kiểm tra kết nối'}</button>
          <button className="btn btn-primary" onClick={save}><Save size={17} /> Lưu cấu hình</button>
        </div>
      </section>
    </AppShell>
  )
}
