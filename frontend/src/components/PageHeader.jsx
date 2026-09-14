export default function PageHeader({ index, title, subtitle }) {
  return (
    <div className="page-header">
      <h1><span>{index} — </span>{title}</h1>
      {subtitle && <p>{subtitle}</p>}
    </div>
  )
}
