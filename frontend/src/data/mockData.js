export const initialProjects = [
  { id: 'website-tmdt', name: 'Website TMĐT', updated: '2 giờ trước' },
  { id: 'ngan-hang', name: 'Ứng dụng ngân hàng', updated: '1 ngày trước' },
  { id: 'crm', name: 'Hệ thống CRM', updated: '3 ngày trước' },
  { id: 'api-thanh-toan', name: 'API Thanh toán', updated: '5 ngày trước' },
  { id: 'noi-bo', name: 'Ứng dụng nội bộ', updated: '1 tuần trước' },
]

export const webFolders = [
  { id: 'dang-nhap', name: 'Đăng nhập', count: 24, updated: 'hôm nay', tone: 'blue' },
  { id: 'khach-hang', name: 'Khách hàng', count: 32, updated: '1 ngày trước', tone: 'green' },
  { id: 'thanh-toan', name: 'Thanh toán', count: 28, updated: '2 ngày trước', tone: 'orange' },
  { id: 'phe-duyet', name: 'Phê duyệt', count: 14, updated: '3 ngày trước', tone: 'purple' },
]

export const apiFolders = [
  { id: 'xac-thuc', name: 'Xác thực', count: 36, updated: 'hôm nay', tone: 'blue' },
  { id: 'khach-hang-api', name: 'Khách hàng', count: 22, updated: '1 ngày trước', tone: 'green' },
  { id: 'thanh-toan-api', name: 'Thanh toán', count: 18, updated: '2 ngày trước', tone: 'orange' },
  { id: 'bao-cao', name: 'Báo cáo', count: 16, updated: '3 ngày trước', tone: 'purple' },
]

export const initialTestcases = [
  {
    id: 'TC-001',
    name: 'Đăng nhập thành công với tài khoản hợp lệ',
    type: 'Chức năng',
    preCondition: 'Người dùng đã có tài khoản hợp lệ và hệ thống đang hoạt động bình thường.',
    testData: 'Email: user@test.com\nMật khẩu: 123456',
    steps: ['Login vào hệ thống', 'Nhập ký tự', 'Quan sát kết quả'],
    expectedResult: 'Đăng nhập thành công và chuyển hướng đến trang chủ, hiển thị tên người dùng trên góc phải màn hình.',
  },
  {
    id: 'TC-002',
    name: 'Hiển thị thông báo lỗi khi nhập sai mật khẩu',
    type: 'Ngoại lệ',
    preCondition: 'Người dùng đã tồn tại trong hệ thống.',
    testData: 'Email: user@test.com\nMật khẩu: sai-mat-khau',
    steps: ['Login vào hệ thống', 'Nhập mật khẩu không đúng', 'Bấm Đăng nhập', 'Quan sát kết quả'],
    expectedResult: 'Hệ thống hiển thị thông báo đăng nhập không thành công.',
  },
  {
    id: 'TC-003',
    name: 'Kiểm tra hiển thị giao diện trang đăng nhập',
    type: 'Giao diện',
    preCondition: 'Truy cập trang đăng nhập.',
    testData: 'Không có',
    steps: ['Mở trang đăng nhập', 'Quan sát bố cục và thành phần giao diện'],
    expectedResult: 'Các thành phần hiển thị đúng vị trí, không vỡ layout.',
  },
  {
    id: 'TC-004',
    name: 'Kiểm tra validation dữ liệu email',
    type: 'Kiểm tra dữ liệu',
    preCondition: 'Đang ở màn hình đăng nhập.',
    testData: 'Email: abc',
    steps: ['Nhập email sai định dạng', 'Nhập mật khẩu hợp lệ', 'Bấm Đăng nhập'],
    expectedResult: 'Hệ thống hiển thị validation email không hợp lệ.',
  },
  {
    id: 'TC-005',
    name: 'Hiển thị popup quên mật khẩu',
    type: 'Popup',
    preCondition: 'Đang ở màn hình đăng nhập.',
    testData: 'Không có',
    steps: ['Bấm Quên mật khẩu', 'Quan sát popup'],
    expectedResult: 'Popup quên mật khẩu hiển thị đúng nội dung.',
  },
  {
    id: 'TC-006',
    name: 'Đăng xuất sau khi đăng nhập thành công',
    type: 'Luồng',
    preCondition: 'Người dùng đã đăng nhập.',
    testData: 'Tài khoản hợp lệ',
    steps: ['Login vào hệ thống', 'Mở menu tài khoản', 'Chọn Đăng xuất', 'Quan sát kết quả'],
    expectedResult: 'Người dùng được đưa về màn hình đăng nhập.',
  },
]

export const apiKeyRows = [
  { id: 1, provider: 'OpenAI', model: 'GPT-4o', key: 'sk-••••••••••x9K2', status: 'Hoạt động' },
  { id: 2, provider: 'Qwen', model: 'Qwen Max', key: 'sk-••••••••••a7H9', status: 'Hết hạn' },
  { id: 3, provider: 'Google', model: 'Gemini 1.5 Pro', key: 'sk-••••••••••m3F8', status: 'Chưa kiểm tra' },
]
