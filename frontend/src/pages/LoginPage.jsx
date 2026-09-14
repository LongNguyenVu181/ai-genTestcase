import { ArrowRight, BarChart3, Building2, CheckCircle2, FileText, Globe2, Grid2X2, LockKeyhole, Mail } from 'lucide-react'
import { useNavigate } from 'react-router-dom'
import Brand from '../components/Brand'

export default function LoginPage() {
  const navigate = useNavigate()
  const submit = (e) => {
    e.preventDefault()
    navigate('/dashboard')
  }

  return (
    <div className="login-page login-page-v2">
      <section className="login-promo">
        <div className="login-promo-inner">
          <Brand />
          <div className="login-index">00 — Đăng nhập</div>

          <div className="login-copy">
            <h1>
              Tạo testcase <span>nhanh hơn,</span>
              <br />
              rõ ràng hơn.
            </h1>
            <p>
              TestPilot AI hỗ trợ phân tích tài liệu và tạo testcase thông minh
              cho Web App và API.
            </p>

            <div className="stat-row">
              <div className="stat-chip">
                <FileText size={22} />
                <div><b>386 testcase</b><span>đã tạo</span></div>
              </div>
              <div className="stat-chip">
                <BarChart3 size={22} />
                <div><b>92% coverage</b><span>trung bình</span></div>
              </div>
              <div className="stat-chip">
                <Grid2X2 size={22} />
                <div><b>12 màn hình</b><span>được phân tích</span></div>
              </div>
            </div>

            <div className="hero-preview" aria-hidden="true">
              <div className="hero-browser">
                <div className="hero-browser-bar">
                  <span></span><span></span><span></span>
                  <div className="hero-browser-title">TestPilot AI</div>
                </div>
                <div className="hero-browser-body">
                  <div className="hero-side">
                    <div className="hero-menu active"></div>
                    <div className="hero-menu"></div>
                    <div className="hero-menu"></div>
                    <div className="hero-menu short"></div>
                  </div>
                  <div className="hero-content">
                    <div className="hero-line hero-line-title"></div>
                    <div className="hero-line"></div>
                    <div className="hero-upload">
                      <FileText size={32} />
                      <b>Kéo thả tài liệu vào đây</b>
                      <span>PDF, MD, DOC, DOCX</span>
                    </div>
                  </div>
                </div>
              </div>

              <div className="analysis-card">
                <div className="analysis-card-title">AI đang phân tích...</div>
                <div><CheckCircle2 size={16} /> Đọc nội dung tài liệu</div>
                <div><CheckCircle2 size={16} /> Trích xuất yêu cầu</div>
                <div className="is-loading"><span className="loading-dot"></span> Tạo testcase</div>
                <div className="analysis-progress"><i></i></div>
              </div>
            </div>
          </div>
        </div>

        <div className="login-promo-footer">
          <span className="footer-accent"></span>
          Cùng QA xây dựng sản phẩm tốt hơn, mỗi ngày.
        </div>
      </section>

      <section className="login-form-area">
        <div className="login-language"><Globe2 size={18} /> VI <span>⌄</span></div>

        <form className="login-card login-card-v2" onSubmit={submit}>
          <Brand compact />
          <h2>Đăng nhập</h2>
          <p>Chào mừng bạn quay trở lại TestPilot AI</p>

          <label className="field-label">Email / User ID</label>
          <div className="input-icon">
            <Mail size={18} />
            <input required placeholder="Nhập email hoặc user ID" />
          </div>

          <label className="field-label">Mật khẩu</label>
          <div className="input-icon">
            <LockKeyhole size={18} />
            <input type="password" required placeholder="Nhập mật khẩu" />
          </div>

          <div className="form-row between">
            <label className="check"><input type="checkbox" defaultChecked /> Ghi nhớ đăng nhập</label>
            <button type="button" className="text-link">Quên mật khẩu?</button>
          </div>

          <button className="btn btn-primary btn-full big login-primary" type="submit">
            Đăng nhập <ArrowRight size={18} />
          </button>

          <div className="divider"><span>hoặc</span></div>

          <button className="btn btn-outline btn-full big login-sso" type="button">
            <Building2 size={18} /> Đăng nhập bằng SSO nội bộ
          </button>

          <div className="login-note">Chưa có tài khoản? Liên hệ quản trị hệ thống để được cấp quyền.</div>
        </form>

        <div className="login-footer-links"><span>Điều khoản</span><i></i><span>Bảo mật</span><i></i><span>Hỗ trợ</span></div>
      </section>
    </div>
  )
}
