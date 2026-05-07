# Zalo Auto-Sender (v1.0.0)

Công cụ tự động gửi tin nhắn và file thông báo qua Zalo Web dựa trên danh sách Excel.

## 🌟 Tính năng chính
- **Tự động hóa Zalo:** Tìm kiếm SĐT, gửi kết bạn và gửi tin nhắn tự động.
- **Đính kèm thông minh:** Tự động tìm và đính kèm file thông báo (.pdf hoặc .png) theo mã căn hộ.
- **Cá nhân hóa tin nhắn:** Hỗ trợ placeholder `{apt}`, `{month_year}`, `{period}`, `{deadline}` để nội dung tin nhắn thay đổi linh hoạt theo từng căn hộ.
- **Phòng tránh Spam:** Kiểm tra lịch sử nhắn tin để tránh gửi trùng lặp tin nhắn hoặc file trong cùng một đợt.
- **Dễ sử dụng:** Giao diện hiện đại, hỗ trợ Dark Mode, cài đặt đơn giản như một ứng dụng Windows thông thường.

## 🛠 Yêu cầu hệ thống
- Hệ điều hành: Windows 10/11.
- Không cần cài đặt Python (nếu dùng bản Setup).
- Kết nối Internet ổn định.

## 🚀 Hướng dẫn sử dụng

### 1. Cài đặt và Đăng nhập
- **Tải bộ cài đặt:** [Tải ZaloAutoSender v1.1.2 (MEGA.nz)](https://mega.nz/file/KkQl0aLR#eXqzEKotP9mBjAjJDuAofFIMj74jk_Mng60XZMhG-8I)
- Chạy file `ZaloAutoSender_Setup.exe` để cài đặt.
- Mở ứng dụng từ Desktop.
- Nhấn **"Đăng nhập (QR)"**, quét mã Zalo trên trình duyệt hiện lên. Khi thấy danh sách tin nhắn hiện ra là đã thành công. Tắt trình duyệt để quay lại App.

### 3. Cấu hình gửi
- **File Excel:** Chọn đường dẫn đến file Excel của bạn.
- **Tên Sheet:** Nhập tên sheet chứa dữ liệu (ví dụ: Sheet1).
- **Cấu hình cột:** Nhập tên cột tương ứng (A, B, C...).
- **Nội dung tin nhắn:** Soạn nội dung, sử dụng các từ khóa `{apt}`... để app tự điền thông tin.
- **Thư mục File:** Chọn thư mục chứa các file thông báo.

### 4. Bắt đầu gửi
- Nhấn **"BẮT ĐẦU GỬI"**. App sẽ mở trình duyệt và tự động thực hiện các thao tác.
- Bạn có thể theo dõi tiến độ ở phần **"Nhật ký hoạt động"**.
- Nhấn **"DỪNG"** bất cứ lúc nào nếu muốn tạm dừng tiến trình.

## 👨‍💻 Thông tin phát triển
- **Phiên bản:** 1.0.0
- **Nguồn:** CMTHANG
- **Công nghệ:** Python, Playwright, CustomTkinter, Pandas.

---
*Lưu ý: Ứng dụng này phục vụ mục đích hỗ trợ nhắc phí, vui lòng không sử dụng để gửi tin nhắn rác (Spam) vi phạm chính sách của Zalo.*
