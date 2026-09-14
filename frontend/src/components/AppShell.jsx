import { Globe2, ChevronDown } from 'lucide-react'
import Sidebar from './Sidebar'

export default function AppShell({ children, wide = false }) {
  return (
    <div className="app-shell">
      <Sidebar />
      <div className="app-main">
        <header className="topbar">
          <div />
          <div className="topbar-actions">
            <button className="topbar-ghost"><Globe2 size={18} /> VI <ChevronDown size={14} /></button>
            <div className="avatar">ND</div>
            <button className="topbar-ghost">Nguyễn Văn A <ChevronDown size={14} /></button>
          </div>
        </header>
        <main className={`content ${wide ? 'content-wide' : ''}`}>{children}</main>
      </div>
    </div>
  )
}
