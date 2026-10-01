# Đưa website lên Internet bằng Render

Bản V5.24.6 đã có sẵn `Dockerfile`, `docker-compose.yml`, `render.yaml`,
Gunicorn và vùng lưu dữ
liệu bền vững. Docker cũng cài LibreOffice để giữ chức năng đọc/chuyển đổi
Word, PDF, Excel và PowerPoint trên máy chủ.

## Cách triển khai

1. Tạo một kho mã riêng tư trên GitHub và đưa toàn bộ thư mục dự án lên đó.
2. Đăng nhập Render, chọn **New + → Blueprint** rồi kết nối kho mã vừa tạo.
3. Render sẽ đọc `render.yaml`. Kiểm tra dịch vụ có ổ đĩa `exam-data` gắn tại
   `/app/storage`; nhập `ADMIN_INITIAL_PASSWORD` và `OPENAI_API_KEY` khi
   Render yêu cầu, sau đó chọn
   **Apply**. Không ghi khóa API trực tiếp vào tệp mã nguồn.
4. Chờ trạng thái **Live** rồi mở địa chỉ dạng
   `https://quan-ly-thi-truc-tuyen-....onrender.com`.

## Việc cần làm ngay sau lần mở đầu tiên

- Đăng nhập quản trị bằng mật khẩu đã nhập ở `ADMIN_INITIAL_PASSWORD`, sau đó
  đổi mật khẩu trong hệ thống nếu cần. Production không tạo tài khoản demo.
- Tải thử một đề DOCX có MathType và một đề PDF có hình.
- Tạo tài khoản thử cho giáo viên/học sinh trước khi dùng chính thức.
- Không xóa ổ đĩa `exam-data`: cơ sở dữ liệu, đề tải lên, hình câu hỏi, tài
  liệu và tin tức đều được lưu tại đó.
- Render chỉ hỗ trợ ổ đĩa bền vững cho dịch vụ trả phí. Không dùng gói miễn
  phí cho dữ liệu thật vì tệp SQLite và tệp giáo viên tải lên có thể mất khi
  máy chủ được tạo lại.
- Chức năng **Tạo đề bằng AI** và **Giáo viên AI** cần `OPENAI_API_KEY`.
  Không ghi khóa API vào mã nguồn hoặc đưa tệp `.env` lên GitHub.

## Chạy thử Docker trên máy tính

1. Sao chép `.env.example` thành `.env` và thay `SECRET_KEY` bằng chuỗi bí mật
   dài, ngẫu nhiên.
2. Chạy `docker compose up --build -d`.
3. Mở `http://localhost:10000`.

Khi dùng tên miền riêng, thêm tên miền trong phần **Custom Domains** của dịch
vụ Render sau khi địa chỉ `onrender.com` đã hoạt động.
