import { ChevronRight } from 'lucide-react'

export default function Breadcrumb({ items }) {
  return (
    <div className="breadcrumb">
      {items.map((item, idx) => (
        <span key={`${item}-${idx}`} className="breadcrumb-part">
          {idx > 0 && <ChevronRight size={15} />}
          <span>{item}</span>
        </span>
      ))}
    </div>
  )
}
