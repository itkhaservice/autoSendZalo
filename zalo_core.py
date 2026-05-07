import asyncio
import pandas as pd
import os
import json
import sys
import traceback
from playwright.async_api import async_playwright

# Cấu hình Playwright cho môi trường EXE
if getattr(sys, 'frozen', False):
    # Ép Playwright sử dụng trình duyệt đã đóng gói trong thư mục ms-playwright
    base_path = sys._MEIPASS
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(base_path, "ms-playwright")
    # Chúng ta KHÔNG xóa PLAYWRIGHT_BROWSERS_PATH ở đây vì cần nó để trỏ vào bundle

def col_to_index(col_str):
    """Chuyển đổi tên cột Excel (A, B, AA...) sang chỉ số 0-based index."""
    exp = 0
    index = 0
    for char in reversed(col_str.upper()):
        index += (ord(char) - ord('A') + 1) * (26 ** exp)
        exp += 1
    return index - 1

class ZaloAutoSender:
    def __init__(self, storage_state="storage_state.json", log_callback=None):
        self.storage_state = storage_state
        self.log_callback = log_callback
        self.is_running = False

    def log(self, message):
        print(message)
        if self.log_callback:
            self.log_callback(message)

    async def login(self):
        self.log("[*] Đang khởi tạo trình duyệt...")
        try:
            async with async_playwright() as p:
                self.log("[*] Đang mở Chromium (Headless=False)...")
                try:
                    browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
                except Exception as launch_err:
                    self.log(f"[!] Lỗi khi mở trình duyệt: {launch_err}")
                    self.log("[!] Gợi ý: Hãy thử chạy 'playwright install chromium' trong CMD.")
                    return False

                # Sử dụng context riêng biệt với storage_state để không ảnh hưởng đến trình duyệt chính hoặc app desktop
                context = await browser.new_context(no_viewport=True)
                page = await context.new_page()
                
                self.log("[*] Đang truy cập chat.zalo.me...")
                await page.goto("https://chat.zalo.me/")
                
                self.log("[!] Vui lòng quét mã QR để đăng nhập.")
                try:
                    await page.wait_for_selector("#contact-search-input", timeout=120000)
                    self.log("[+] Đăng nhập thành công!")
                    await asyncio.sleep(5)
                    await context.storage_state(path=self.storage_state)
                    self.log(f"[+] Đã lưu phiên làm việc vào {self.storage_state}")
                    return True
                except Exception as e:
                    self.log(f"[!] Lỗi hoặc quá thời gian đăng nhập: {e}")
                    return False
                finally:
                    await browser.close()
        except Exception as p_err:
            self.log(f"[!] Lỗi Playwright: {p_err}")
            self.log(traceback.format_exc())
            return False

    async def check_login(self):
        if not os.path.exists(self.storage_state):
            self.log(f"[!] Không tìm thấy {self.storage_state}")
            return False

        self.log("[*] Đang kiểm tra trạng thái đăng nhập...")
        try:
            async with async_playwright() as p:
                # Thêm args=["--start-maximized"] và no_viewport=True để đồng nhất giao diện
                browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
                context = await browser.new_context(storage_state=self.storage_state, no_viewport=True)
                page = await context.new_page()
                await page.goto("https://chat.zalo.me/")
                
                try:
                    await page.wait_for_selector("#contact-search-input", timeout=15000)
                    self.log("[+] Phiên đăng nhập vẫn còn hiệu lực.")
                    await context.storage_state(path=self.storage_state) # Update cookies
                    return True
                except:
                    self.log("[!] Phiên làm việc đã hết hạn hoặc không load được.")
                    return False
                finally:
                    await browser.close()
        except Exception as e:
            self.log(f"[!] Lỗi khi kiểm tra session: {e}")
            return False

    async def send_process(self, excel_path, sheet_name, file_dir, message_template, config):
        if not os.path.exists(self.storage_state):
            self.log("[!] Cần đăng nhập trước khi gửi.")
            return

        self.is_running = True
        try:
            self.log(f"[*] Đang đọc file Excel: {excel_path}")
            if not os.path.exists(excel_path):
                self.log(f"[!] File Excel không tồn tại: {excel_path}")
                return

            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
            
            # Chuyển đổi tên cột chữ cái sang index
            try:
                col_apt_idx = col_to_index(config.get("col_apt", "B"))
                col_sdt_idx = col_to_index(config.get("col_sdt", "J"))
            except Exception as col_err:
                self.log(f"[!] Lỗi cấu hình cột: {col_err}")
                return
            
            num_cols = len(df.columns)
            max_idx = max(col_apt_idx, col_sdt_idx)
            if num_cols <= max_idx:
                self.log(f"[!] Lỗi: File Excel chỉ có {num_cols} cột. Không thể truy cập cột chỉ định.")
                return

            data = []
            for index, row in df.iterrows():
                if not self.is_running: break
                try:
                    apt = str(row.iloc[col_apt_idx]).strip()
                    sdt = str(row.iloc[col_sdt_idx]).strip()
                    
                    if sdt.endswith('.0'): sdt = sdt[:-2]
                    if sdt and sdt != "nan" and apt and apt != "nan":
                        sdt = sdt.replace(" ", "")
                        if not sdt.startswith('0') and len(sdt) >= 9 and sdt.isdigit():
                            sdt = '0' + sdt
                        data.append({"apt": apt, "sdt": sdt})
                except:
                    continue
            
            if not data:
                self.log("[!] Không tìm thấy dữ liệu hợp lệ.")
                return

            self.log(f"[+] Tìm thấy {len(data)} khách hàng cần gửi.")
            
            processed_in_session = set()

            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False, args=["--start-maximized"])
                context = await browser.new_context(storage_state=self.storage_state, no_viewport=True)
                page = await context.new_page()
                page.set_default_timeout(20000)
                
                await page.goto("https://chat.zalo.me/")
                await page.wait_for_selector("#contact-search-input")

                for i, item in enumerate(data):
                    if not self.is_running: 
                        self.log("[!] Đã dừng tiến trình theo yêu cầu.")
                        break
                        
                    apt = item["apt"]
                    sdt = item["sdt"]
                    
                    if apt in processed_in_session:
                        continue

                    # Kiểm tra đính kèm file
                    should_attach = config.get("attach_file", True)
                    file_path = ""
                    if should_attach:
                        file_path = os.path.normpath(os.path.join(file_dir, f"{apt}.pdf"))
                        if not os.path.exists(file_path):
                            file_path = os.path.normpath(os.path.join(file_dir, f"{apt}.png"))
                        
                        if not os.path.exists(file_path):
                            self.log(f"[-] [{i+1}/{len(data)}] Thiếu file cho căn {apt}. Bỏ qua.")
                            continue

                    self.log(f"[>] [{i+1}/{len(data)}] Đang gửi: {apt} ({sdt})")
                    
                    try:
                        search_input = await page.wait_for_selector("#contact-search-input")
                        await search_input.click()
                        await page.keyboard.press("Control+A")
                        await page.keyboard.press("Backspace")
                        await search_input.fill(sdt)
                        
                        try:
                            await page.wait_for_selector(".conv-item", timeout=4000)
                        except:
                            self.log(f"[!] Không tìm thấy Zalo cho sđt {sdt}")
                            continue

                        await page.locator(".conv-item").first.click()
                        await asyncio.sleep(2)

                        try:
                            history_content = await page.locator(".message-view").inner_text()
                        except:
                            history_content = ""
                            
                        # Tạo tin nhắn từ template
                        msg = message_template.replace("{apt}", apt)
                        msg = msg.replace("{month_year}", config.get("month_year", ""))
                        msg = msg.replace("{period}", config.get("period", ""))
                        msg = msg.replace("{deadline}", config.get("deadline", ""))

                        check_keyword = msg.split('\n')[0][:20]
                        
                        # Kiểm tra trùng tin nhắn hoặc trùng tên file
                        is_duplicate = False
                        if check_keyword and check_keyword in history_content:
                            is_duplicate = True
                        if should_attach and os.path.basename(file_path) in history_content:
                            is_duplicate = True
                            
                        if is_duplicate:
                            self.log(f"[=] {apt} đã được gửi trước đó. Bỏ qua.")
                            processed_in_session.add(apt)
                            continue

                        btn_kb = page.locator("[data-translate-inner='STR_FRIEND_REQ_SEND']").first
                        if await btn_kb.is_visible():
                            await btn_kb.click()
                            await asyncio.sleep(1)
                            confirm = page.locator(".btn-primary:has-text('Gửi lời mời'), .btn-primary:has-text('Kết bạn')").last
                            if await confirm.is_visible():
                                await confirm.click()

                        chat_input = page.locator("#richInput")
                        if await chat_input.is_visible():
                            await chat_input.focus()
                            await page.keyboard.type(msg)
                            await page.keyboard.press("Enter")
                            await asyncio.sleep(1)
                        
                            if should_attach and file_path:
                                attach_btn = page.locator("div[title='Đính kèm File'], [icon='Attach_24_Line']").first
                                if await attach_btn.is_visible():
                                    await attach_btn.click()
                                    await asyncio.sleep(1)
                                    async with page.expect_file_chooser() as fc_info:
                                        await page.locator(".zmenu-item[data-id='div_CX_Select'], [data-translate-inner='STR_CHOOSE_FILE_COMPUTER']").first.click()
                                    file_chooser = await fc_info.value
                                    await file_chooser.set_files(file_path)
                                    await asyncio.sleep(4)
                            
                            self.log(f"[+] Gửi thành công cho {apt}")
                            processed_in_session.add(apt)
                            await asyncio.sleep(2)
                                
                    except Exception as e:
                        self.log(f"[!] Lỗi khi xử lý {apt}: {e}")
                        continue
                
                await browser.close()
        except Exception as e:
            self.log(f"[!] Lỗi hệ thống: {e}")
        finally:
            self.is_running = False
            self.log("--- TIẾN TRÌNH KẾT THÚC ---")

    def stop(self):
        self.is_running = False
