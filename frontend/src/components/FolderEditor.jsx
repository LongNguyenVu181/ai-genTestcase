import { GripVertical, Plus, X } from 'lucide-react'
import { useState } from 'react'

export default function FolderEditor({ title, icon, defaultItems = [], onChange }) {
  const [items, setItems] = useState(defaultItems)
  const [draft, setDraft] = useState('')

  const commit = next => {
    setItems(next)
    onChange?.(next)
  }

  const add = () => {
    const name = draft.trim() || `Thư mục ${items.length + 1}`
    commit([...items, name])
    setDraft('')
  }

  return (
    <div className="folder-editor">
      <div className="folder-editor-title">{icon}{title}</div>
      <div className="folder-editor-list">
        {items.map((item, index) => (
          <div className="folder-editor-item" key={`${item}-${index}`}>
            <GripVertical size={16} />
            <span>{item}</span>
            <button className="icon-btn small" onClick={() => commit(items.filter((_, i) => i !== index))}><X size={16} /></button>
          </div>
        ))}
      </div>
      <div className="folder-add-row">
        <input value={draft} onChange={e => setDraft(e.target.value)} placeholder="Tên thư mục..." onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), add())} />
        <button className="folder-add" type="button" onClick={add}><Plus size={16} /> Thêm thư mục</button>
      </div>
    </div>
  )
}
