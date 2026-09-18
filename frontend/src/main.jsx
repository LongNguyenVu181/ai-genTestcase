import React from 'react'
import ReactDOM from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import App from './App'
import './styles.css'

// The application is entirely browser-session-scoped. A refresh starts with an
// empty workspace, including AI keys, files, projects, runs and testcases.
const clearTransientSession = () => {
  try {
    for (const key of Object.keys(sessionStorage)) {
      if (key.startsWith('testpilot.')) sessionStorage.removeItem(key)
    }
  } catch {
    // The app remains usable when browser storage is unavailable.
  }
}

clearTransientSession()

// Clear immediately before navigation/reload as well as at next bootstrap.
window.addEventListener('pagehide', clearTransientSession, { once: true })

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <BrowserRouter>
      <App />
    </BrowserRouter>
  </React.StrictMode>,
)
