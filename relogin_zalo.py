import asyncio
import os
from playwright.async_api import async_playwright

STORAGE_STATE = "storage_state.json"

async def relogin():
    if not os.path.exists(STORAGE_STATE):
        print(f"Không tìm thấy {STORAGE_STATE}. Vui lòng chạy login_zalo.py trước.")
        return

    async with async_playwright() as p:
        # Khởi chạy trình duyệt
        browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
        
        # Tạo context từ session đã lưu
        context = await browser.new_context(storage_state=STORAGE_STATE, no_viewport=True)
        page = await context.new_page()
        
        print(f"[*] Đang tải phiên làm việc từ {STORAGE_STATE}...")
        await page.goto("https://chat.zalo.me/")
        
        try:
            # Chờ xem có vào được trang chủ chat không (timeout 20s)
            await page.wait_for_selector("#contact-search-input", timeout=20000)
            print("[+] Đăng nhập lại thành công!")
            
            print("[!] Bạn có thể kiểm tra Zalo ngay bây giờ.")
            print("[!] Khi xong, hãy nhấn Enter tại cửa sổ lệnh này để lưu session mới và đóng trình duyệt.")
            
            # Chờ người dùng nhấn Enter để cập nhật session mới nhất
            input(">>> NHẤN ENTER ĐỂ LƯU VÀ THOÁT...")
            
            # Lưu lại session mới nhất (phòng trường hợp cookies được update)
            await context.storage_state(path=STORAGE_STATE)
            print(f"[+] Đã cập nhật session mới nhất vào {STORAGE_STATE}")
            
        except Exception as e:
            print(f"[!] Phiên làm việc đã hết hạn hoặc có lỗi: {e}")
            print("[!] Vui lòng chạy login_zalo.py để quét lại mã QR.")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(relogin())
