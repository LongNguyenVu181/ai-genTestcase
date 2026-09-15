import { Navigate, Route, Routes } from 'react-router-dom'
import LoginPage from './pages/LoginPage'
import DashboardPage from './pages/DashboardPage'
import ProjectDetailPage from './pages/ProjectDetailPage'
import ScopeExplorerPage from './pages/ScopeExplorerPage'
import UploadPage from './pages/UploadPage'
import TestcasePage from './pages/TestcasePage'
import ApiKeysPage from './pages/ApiKeysPage'

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Navigate to="/login" replace />} />
      <Route path="/login" element={<LoginPage />} />
      <Route path="/dashboard" element={<DashboardPage />} />
      <Route path="/project/:projectId" element={<ProjectDetailPage />} />

      {/* Project -> Web/API -> Screen -> Testcase */}
      <Route path="/project/:projectId/scope/:scope" element={<ScopeExplorerPage />} />
      <Route path="/project/:projectId/scope/:scope/upload" element={<UploadPage />} />
      <Route path="/project/:projectId/scope/:scope/testcases" element={<TestcasePage />} />

      {/* Custom workspace routes */}
      <Route path="/project/:projectId/folder/:folderId/upload" element={<UploadPage />} />
      <Route path="/project/:projectId/folder/:folderId/testcases" element={<TestcasePage />} />

      <Route path="/api-keys" element={<ApiKeysPage />} />
      <Route path="*" element={<Navigate to="/dashboard" replace />} />
    </Routes>
  )
}
