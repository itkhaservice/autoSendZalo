import asyncio
import pandas as pd
import os
from playwright.async_api import async_playwright

# Cấu hình đường dẫn
EXCEL_FILE = "PHAI_THU_CU_DAN THÁNG 4.xlsx"
SHEET_NAME = "04.2026"
PDF_DIR = r"C:\Users\qlnha\Desktop\RIVA\Auto-send-zalo\NHAC NO THANG 04\NHAC NO THANG 04"
STORAGE_STATE = "storage_state.json"

async def send_files():
    if not os.path.exists(STORAGE_STATE):
        print(f"Không tìm thấy {STORAGE_STATE}. Vui lòng chạy login_zalo.py trước.")
        return

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
        data = []
        for index, row in df.iterrows():
            apt = str(row.iloc[1]).strip()
            sdt = str(row.iloc[9]).strip()
            if sdt.endswith('.0'): sdt = sdt[:-2]
            if sdt and sdt != "nan" and apt and apt != "nan":
                sdt = sdt.replace(" ", "")
                if not sdt.startswith('0') and len(sdt) >= 9 and sdt.isdigit():
                    sdt = '0' + sdt
                data.append({"apt": apt, "sdt": sdt})
    except Exception as e:
        print(f"Lỗi khi đọc file Excel: {e}")
        return

    print(f"Tìm thấy {len(data)} khách hàng cần gửi.")
    
    # Tập hợp các căn hộ đã xử lý trong phiên này để tránh trùng lặp nếu Excel bị lặp dòng
    processed_in_session = set()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
        context = await browser.new_context(storage_state=STORAGE_STATE, no_viewport=True)
        page = await context.new_page()
        
        page.set_default_timeout(15000)
        await page.goto("https://chat.zalo.me/")
        
        try:
            await page.wait_for_selector("#contact-search-input")
        except:
            print("Không load được trang Zalo.")
            return

        for i, item in enumerate(data):
            apt = item["apt"]
            sdt = item["sdt"]
            
            if apt in processed_in_session:
                print(f"[-] [{i+1}/{len(data)}] {apt} đã được xử lý trong lượt này. Bỏ qua.")
                continue

            pdf_path = os.path.normpath(os.path.join(PDF_DIR, f"{apt}.pdf"))
            if not os.path.exists(pdf_path):
                print(f"[-] [{i+1}/{len(data)}] Thiếu file: {apt}")
                continue

            print(f"[>] [{i+1}/{len(data)}] Xử lý: {apt} - {sdt}")
            
            try:
                # 1. Tìm kiếm
                search_input = await page.wait_for_selector("#contact-search-input")
                await search_input.click()
                await page.keyboard.press("Control+A")
                await page.keyboard.press("Backspace")
                await search_input.fill(sdt)
                
                try:
                    await page.wait_for_selector(".conv-item", timeout=3000)
                except:
                    print(f"[!] Không thấy Zalo cho {sdt}")
                    continue

                await page.locator(".conv-item").first.click()
                # Tăng thời gian chờ load chat để check lịch sử chính xác hơn
                await asyncio.sleep(3) 

                # 2. Gửi kết bạn (nếu có nút trên banner)
                btn_kb = page.locator("[data-translate-inner='STR_FRIEND_REQ_SEND']").first
                if await btn_kb.is_visible():
                    print(f"[*] Đang bấm Gửi kết bạn cho {sdt}...")
                    await btn_kb.click()
                    await asyncio.sleep(1.5)
                    confirm = page.locator(".btn-primary:has-text('Gửi lời mời'), .btn-primary:has-text('Kết bạn')").last
                    if await confirm.is_visible():
                        await confirm.click()
                        await asyncio.sleep(1)

                # 3. KIỂM TRA LỊCH SỬ CHẶN GỬI TRÙNG
                # Quét nội dung chat tìm keyword hoặc tên file PDF
                history_content = await page.locator(".message-view").inner_text()
                has_keyword = "thông báo phí tháng 04/2026" in history_content
                has_file = f"{apt}.pdf" in history_content
                
                if has_keyword or has_file:
                    print(f"[=] Đã có lịch sử gửi cho {apt} (Keyword: {has_keyword}, File: {has_file}). DỪNG GỬI.")
                    processed_in_session.add(apt)
                    continue

                # 4. Gửi lời nhắn
                chat_input = page.locator("#richInput")
                if await chat_input.is_visible():
                    await chat_input.focus()
                    greeting = (
                        f"Chào anh/chị căn hộ {apt}, em gửi thông báo phí tháng 04/2026 ạ. "
                        "Đây là tin nhắn tự động nhắc đóng phí đợt 2 (hạn chót đến ngày 28/04/2026). "
                        "Anh/chị vui lòng bỏ qua thông báo này nếu đã hoàn tất thanh toán. Em xin cảm ơn!"
                    )
                    await page.keyboard.type(greeting)
                    await page.keyboard.press("Enter")
                    await asyncio.sleep(1.5)
                
                    # 5. Gửi File PDF
                    attach_btn = page.locator("div[title='Đính kèm File'], [icon='Attach_24_Line']").first
                    if await attach_btn.is_visible():
                        await attach_btn.click()
                        await asyncio.sleep(1)
                        
                        try:
                            async with page.expect_file_chooser() as fc_info:
                                await page.locator(".zmenu-item[data-id='div_CX_Select'], [data-translate-inner='STR_CHOOSE_FILE_COMPUTER']").first.click()
                            
                            file_chooser = await fc_info.value
                            await file_chooser.set_files(pdf_path)
                            print(f"[+] Gửi thành công cho {apt}")
                            processed_in_session.add(apt)
                            await asyncio.sleep(5) # Đợi upload xong mới sang người tiếp theo
                        except Exception as file_err:
                            print(f"[!] Lỗi chọn file: {file_err}")
                
            except Exception as e:
                print(f"[!] Lỗi khi xử lý {apt}: {e}")
                continue

        print("\n--- HOÀN THÀNH ---")
        await asyncio.sleep(5)
        await browser.close()

if __name__ == "__main__":
    asyncio.run(send_files())
