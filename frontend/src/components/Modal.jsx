import { X } from 'lucide-react'

export default function Modal({ open, title, children, onClose, width = '560px' }) {
  if (!open) return null
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal" style={{ maxWidth: width }} onMouseDown={e => e.stopPropagation()}>
        <div className="modal-head">
          <h3>{title}</h3>
          <button className="icon-btn" onClick={onClose}><X size={20} /></button>
        </div>
        {children}
      </div>
    </div>
  )
}
