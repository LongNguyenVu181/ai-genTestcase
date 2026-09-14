export default function Brand({ compact = false }) {
  return (
    <div className={`brand ${compact ? 'brand--compact' : ''}`}>
      <span className="brand-mark" aria-hidden="true">
        <i></i><i></i><i></i>
      </span>
      <span className="brand-name">TestPilot <b>AI</b></span>
    </div>
  )
}
