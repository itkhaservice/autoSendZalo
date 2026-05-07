import asyncio
from playwright.async_api import async_playwright

async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("https://chat.zalo.me/")
        
        print("Vui lòng quét mã QR để đăng nhập Zalo Web.")
        print("Sau khi đăng nhập xong, hãy đợi 1 chút để hệ thống lưu phiên làm việc.")
        
        # Đợi cho đến khi xuất hiện phần tử đặc trưng của trang đã đăng nhập
        # Ví dụ: thanh tìm kiếm hoặc danh sách bạn bè
        try:
            await page.wait_for_selector("#contact-search-input", timeout=120000) # Đợi tối đa 2 phút
            print("Đăng nhập thành công!")
            await asyncio.sleep(5) # Đợi thêm vài giây để đồng bộ dữ liệu
            await context.storage_state(path="storage_state.json")
            print("Đã lưu phiên làm việc vào storage_state.json")
        except Exception as e:
            print(f"Lỗi hoặc quá thời gian đăng nhập: {e}")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(run())
