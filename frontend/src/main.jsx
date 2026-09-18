import React from 'react'
import ReactDOM from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import App from './App'
import './styles.css'

const SESSION_ID_KEY = 'testpilot.browser-session.v1'

// The application is entirely session-scoped. Notify the backend first so its
// matching in-memory database is discarded, then remove browser session data.
const clearTransientSession = () => {
  try {
    const previousSessionId = sessionStorage.getItem(SESSION_ID_KEY)
    if (previousSessionId && navigator.sendBeacon) {
      navigator.sendBeacon(`/api/session/end?session_id=${encodeURIComponent(previousSessionId)}`)
    }
    for (const key of Object.keys(sessionStorage)) {
      if (key.startsWith('testpilot.')) sessionStorage.removeItem(key)
    }
  } catch {
    // The app remains usable when browser storage is unavailable.
  }
}

clearTransientSession()

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <BrowserRouter>
      <App />
    </BrowserRouter>
  </React.StrictMode>,
)
