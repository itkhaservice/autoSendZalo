# Zalo & Email Auto-Sender (v2.0.0)

Công cụ tự động gửi tin nhắn Zalo và Email thông báo qua Webmail (Mắt Bão/Roundcube) dựa trên danh sách Excel.

## 🌟 Tính năng chính
- **Tự động hóa Zalo:** Tìm kiếm SĐT, gửi kết bạn và gửi tin nhắn tự động.
- **Tự động hóa Email:** Gửi email hàng loạt qua giao diện Webmail (Mắt Bão) với nội dung cá nhân hóa.
- **Đính kèm thông minh:** Tự động tìm và đính kèm file thông báo (.pdf hoặc .png) theo mã căn hộ.
- **Cá nhân hóa nội dung:** Hỗ trợ placeholder `{apt}`, `{month_year}`, `{period}`, `{deadline}` để nội dung tin nhắn/email thay đổi linh hoạt.
- **Phòng tránh Spam:** Kiểm tra lịch sử nhắn tin Zalo để tránh gửi trùng lặp.
- **Dễ sử dụng:** Giao diện hiện đại, hỗ trợ Dark Mode, cài đặt đơn giản.

## 🛠 Yêu cầu hệ thống
- Hệ điều hành: Windows 10/11.
- Không cần cài đặt Python (nếu dùng bản Setup).
- Kết nối Internet ổn định.

## 🚀 Hướng dẫn sử dụng

### 1. Cài đặt và Đăng nhập
- **Tải bộ cài đặt:** [Tải Zalo & Email Auto-Sender v2.0.0 (MEGA.nz)](https://mega.nz/file/6l5VCAxZ#qegapNOoJN4Z2v9QJqh9nRkbeNOxrMJqmfh0W1ZHH7s)
- Chạy file `ZaloAutoSender_Setup.exe` để cài đặt.
- Mở ứng dụng từ Desktop.
- **Zalo:** Nhấn "Đăng nhập (QR)", quét mã Zalo trên trình duyệt.
- **Email:** Nhập thông tin Webmail (URL, User, Pass) và nhấn "Đăng nhập Webmail".

### 2. Cấu hình gửi
- **File Excel:** Chọn đường dẫn đến file Excel chứa danh sách căn hộ.
- **Tên Sheet:** Nhập tên sheet chính xác trong file Excel.
- **Cấu hình cột:** Nhập tên cột (A, B, C...) tương ứng với Mã căn hộ, SĐT hoặc Email.
- **Thư mục File:** Chọn thư mục chứa các file thông báo (tên file phải trùng với mã căn hộ).

### 3. Bắt đầu gửi
- Chọn tab **GỬI ZALO** hoặc **GỬI EMAIL**.
- Nhấn **"BẮT ĐẦU GỬI"**. App sẽ mở trình duyệt và thực hiện tự động.
- Theo dõi tiến độ ở phần **"Nhật ký hoạt động"**.
- Nhấn **"DỪNG"** để tạm dừng nếu cần.

## 👨‍💻 Thông tin phát triển
- **Phiên bản:** 2.0.0
- **Nguồn:** CMTHANG
- **Công nghệ:** Python, Playwright, CustomTkinter, Pandas.

---
*Lưu ý: Ứng dụng này phục vụ mục đích hỗ trợ nhắc phí, vui lòng không sử dụng để gửi spam vi phạm chính sách của nhà cung cấp dịch vụ.*

<p align="right"><sub><i>Sản phẩm dành tặng cho Thanh Ngân Riva Park</i></sub></p>
