import asyncio
import pandas as pd
import os
import json
import sys
import traceback
from playwright.async_api import async_playwright

# Cấu hình Playwright cho môi trường EXE
if getattr(sys, 'frozen', False):
    internal_path = os.path.join(sys._MEIPASS, "pw-browsers")
    root_path = os.path.join(os.path.dirname(sys.executable), "pw-browsers")
    if os.path.exists(internal_path):
        browsers_dir = internal_path
    else:
        browsers_dir = root_path
else:
    browsers_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pw-browsers")

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browsers_dir

portable_browser = os.path.join(browsers_dir, "chromium-1208", "chrome-win64", "chrome.exe")
if os.path.exists(portable_browser):
    executable_path = portable_browser
else:
    executable_path = None

def col_to_index(col_str):
    exp = 0
    index = 0
    for char in reversed(col_str.upper()):
        index += (ord(char) - ord('A') + 1) * (26 ** exp)
        exp += 1
    return index - 1

class EmailAutoSender:
    def __init__(self, storage_state="email_storage_state.json", log_callback=None):
        self.storage_state = storage_state
        self.log_callback = log_callback
        self.is_running = False

    def log(self, message):
        print(message)
        if self.log_callback:
            self.log_callback(message)

    async def login(self, url, username, password):
        self.log("[*] Đang khởi tạo trình duyệt Webmail...")
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    executable_path=executable_path,
                    headless=False, 
                    args=["--start-maximized"]
                )
                context = await browser.new_context(no_viewport=True)
                page = await context.new_page()
                
                self.log(f"[*] Đang truy cập {url}...")
                await page.goto(url)
                
                try:
                    # Chờ form đăng nhập cPanel
                    await page.wait_for_selector("#user", timeout=10000)
                    await page.fill("#user", username)
                    await page.fill("#pass", password)
                    await page.click("#login_submit")
                    
                    # Chờ chuyển hướng. Có thể qua trang chọn Webmail hoặc vào thẳng Roundcube
                    self.log("[*] Đang chờ đăng nhập...")
                    
                    # Nếu hiện trang chọn Webmail (cPanel), nhấn vào Roundcube
                    try:
                        roundcube_link = page.locator("a:has-text('Roundcube')").first
                        await roundcube_link.wait_for(state="visible", timeout=10000)
                        await roundcube_link.click()
                    except:
                        pass # Có thể đã vào thẳng

                    # Đợi cho đến khi vào được giao diện chính của Roundcube
                    await page.wait_for_selector(".task-mail, #rcmbtn115, #logo", timeout=30000)
                    
                    # Lưu lại URL thực tế (bao gồm cả cpsess token)
                    current_url = page.url
                    self.log(f"[+] Đã vào Webmail: {current_url}")
                    
                    self.log("[+] Đăng nhập Webmail thành công!")
                    await asyncio.sleep(2)
                    
                    # Lưu state bao gồm cả URL gốc để dùng lại
                    state = await context.storage_state()
                    # Thêm thông tin URL vào metadata nếu cần, hoặc lưu riêng
                    with open(self.storage_state, "w", encoding="utf-8") as f:
                        json.dump(state, f)
                    
                    # Lưu URL cuối cùng để biết đường dẫn gốc của Roundcube
                    url_info_file = self.storage_state + ".url"
                    with open(url_info_file, "w", encoding="utf-8") as f:
                        f.write(current_url)

                    return True
                except Exception as e:
                    self.log(f"[!] Lỗi đăng nhập Webmail: {e}")
                    return False
                finally:
                    await browser.close()
        except Exception as p_err:
            self.log(f"[!] Lỗi Playwright: {p_err}")
            return False

    async def send_process(self, excel_path, sheet_name, file_dir, email_url, subject_template, body_template, config):
        if not os.path.exists(self.storage_state):
            self.log("[!] Cần đăng nhập Webmail trước khi gửi.")
            return

        # Đọc URL thực tế từ file lưu trữ
        url_info_file = self.storage_state + ".url"
        actual_url = email_url
        if os.path.exists(url_info_file):
            with open(url_info_file, "r", encoding="utf-8") as f:
                actual_url = f.read().strip()

        self.is_running = True
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
            col_apt_idx = col_to_index(config.get("col_apt", "B"))
            col_email_idx = col_to_index(config.get("col_email", "K"))

            data = []
            for index, row in df.iterrows():
                if not self.is_running: break
                try:
                    apt = str(row.iloc[col_apt_idx]).strip()
                    email = str(row.iloc[col_email_idx]).strip()
                    if email and email != "nan" and "@" in email:
                        data.append({"apt": apt, "email": email})
                except: continue

            if not data:
                self.log("[!] Không tìm thấy dữ liệu Email hợp lệ.")
                return

            self.log(f"[+] Tìm thấy {len(data)} email cần gửi.")

            # Lấy các tham số chung từ config
            month_year = config.get("month_year", "")
            period = config.get("period", "")
            deadline = config.get("deadline", "")

            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    executable_path=executable_path,
                    headless=False, 
                    args=["--start-maximized"]
                )
                context = await browser.new_context(storage_state=self.storage_state, no_viewport=True)
                page = await context.new_page()
                page.set_default_timeout(30000)
                
                # Truy cập Webmail bằng URL đã lưu (chứa session token)
                self.log(f"[*] Đang mở Webmail: {actual_url}")
                try:
                    await page.goto(actual_url, wait_until="networkidle", timeout=45000)
                except:
                    self.log("[!] Mạng chậm, vẫn tiếp tục quy trình...")
                    await page.goto(email_url, wait_until="domcontentloaded")

                for i, item in enumerate(data):
                    if not self.is_running: break
                    
                    apt = item["apt"]
                    email = item["email"]
                    
                    file_path = os.path.normpath(os.path.join(file_dir, f"{apt}.pdf"))
                    if not os.path.exists(file_path):
                        file_path = os.path.normpath(os.path.join(file_dir, f"{apt}.png"))
                    
                    if not os.path.exists(file_path):
                        self.log(f"[-] [{i+1}/{len(data)}] Thiếu file cho {apt}. Bỏ qua.")
                        continue

                    self.log(f"[>] [{i+1}/{len(data)}] Xử lý: {apt} ({email})")
                    
                    try:
                        # 1. Bấm Soạn thư - Tìm theo văn bản "Soạn thư" hoặc "Compose"
                        self.log("[*] Đang tìm nút Soạn thư...")
                        compose_selectors = [
                            page.get_by_role("link", name="Soạn thư"),
                            page.get_by_role("button", name="Soạn thư"),
                            page.locator("a.compose"),
                            page.locator("[data-command='compose']"),
                            page.locator("#rcmbtn104") # Dự phòng ID phổ biến
                        ]
                        
                        btn_found = None
                        for selector in compose_selectors:
                            try:
                                if await selector.first.is_visible(timeout=2000):
                                    btn_found = selector.first
                                    break
                            except: continue
                        
                        if btn_found:
                            self.log("[*] Đã thấy nút Soạn thư, đang click...")
                            await btn_found.click()
                        else:
                            self.log("[!] Không thấy nút Soạn thư. Thử click cưỡng bức bằng class...")
                            await page.locator("a.compose").first.click(force=True)
                        
                        # Đợi trang soạn thư load xong các thành phần quan trọng
                        self.log("[*] Đang đợi form soạn thư hiển thị...")
                        await page.wait_for_load_state("networkidle")
                        
                        # 2. Nhập người nhận
                        self.log(f"[*] Nhập người nhận: {email}")
                        # Tìm ô input nằm trong vùng nhập người nhận (To)
                        # Elastic skin: Thường là input bên trong .recipient-input
                        to_area = page.locator(".recipient-input, #compose_to").first
                        await to_area.scroll_into_view_if_needed()
                        await to_area.click() # Click vào vùng chứa để kích hoạt input
                        
                        # Tìm ô input thực sự (có role combobox hoặc nằm trong ul)
                        to_input = to_area.locator("input").first
                        await to_input.wait_for(state="visible", timeout=10000)
                        await to_input.press_sequentially(email, delay=50)
                        await asyncio.sleep(0.5)
                        await page.keyboard.press("Enter")
                        await asyncio.sleep(0.5)

                        # 3. Nhập Tiêu đề
                        self.log(f"[*] Nhập tiêu đề cho {apt}")
                        subject = subject_template.replace("{apt}", apt).replace("{month_year}", month_year).replace("{period}", period).replace("{deadline}", deadline)
                        subject_input = page.locator("#compose-subject, input[name='_subject']").first
                        await subject_input.fill(subject)
                        await asyncio.sleep(0.5)

                        # 4. Nhập Nội dung
                        self.log("[*] Nhập nội dung thư...")
                        body = body_template.replace("{apt}", apt).replace("{month_year}", month_year).replace("{period}", period).replace("{deadline}", deadline)
                        
                        # Kiểm tra xem có TinyMCE không
                        editor_iframe = page.frame_locator("iframe[id^='composebody_'], iframe.tox-edit-area__iframe").first
                        try:
                            # Đợi iframe load
                            await asyncio.sleep(1) 
                            if await editor_iframe.locator("body").is_visible(timeout=3000):
                                self.log("[*] Điền vào Rich Text Editor...")
                                await editor_iframe.locator("body").click()
                                await editor_iframe.locator("body").evaluate("(el, text) => { el.innerHTML = text; }", body.replace("\n", "<br>"))
                            else:
                                raise Exception("No Rich Editor")
                        except:
                            # Plain text
                            self.log("[*] Điền vào Plain Text Editor...")
                            await page.locator("#composebody, textarea[name='_message']").first.fill(body)

                        await asyncio.sleep(1)

                        # 5. Kèm tệp tin đính kèm
                        self.log(f"[*] Đang đính kèm tệp: {os.path.basename(file_path)}")
                        async with page.expect_file_chooser() as fc_info:
                            attach_btn = page.locator("[data-command='attach'], #rcmbtn112, .attach, a.button.attach").first
                            await attach_btn.click()
                        file_chooser = await fc_info.value
                        await file_chooser.set_files(file_path)
                        
                        self.log("[*] Đang chờ tệp tải lên hoàn tất...")
                        # Chờ thanh tiến trình biến mất và file hiện lên
                        try:
                            await page.wait_for_selector(".attachmentslist li:not(.uploading), #attachment-list li:not(.uploading)", timeout=45000)
                            self.log("[+] Đã tải tệp lên xong.")
                        except:
                            self.log("[!] Cảnh báo: Tệp tải lên chưa xác nhận, vẫn tiếp tục.")
                        
                        await asyncio.sleep(1)

                        # 6. Gửi
                        self.log("[*] Đang nhấn nút Gửi...")
                        send_btn = page.locator("[data-command='send'], #rcmbtn115, .send, button.send").first
                        await send_btn.click()
                        
                        # Đợi thông báo thành công
                        self.log("[*] Đang đợi hệ thống xác nhận...")
                        try:
                            # Thông báo thành công hoặc quay lại trang Inbox
                            success_indicator = page.locator(".messagemain.success, .notifier.success, #messagestack .success, .task-mail:not(.action-compose)")
                            await success_indicator.first.wait_for(state="visible", timeout=30000)
                            self.log(f"[+] Đã gửi thành công cho {apt}")
                        except:
                            self.log(f"[*] Đã thực hiện lệnh gửi cho {apt}. Hãy kiểm tra Sent box.")

                        await asyncio.sleep(3) 

                    except Exception as e:
                        self.log(f"[!] Lỗi tại {apt}: {e}")
                        # Quay lại trang chính để reset nếu lỗi
                        await page.goto(actual_url, wait_until="domcontentloaded")
                        await asyncio.sleep(2)
                        continue

                    except Exception as e:
                        self.log(f"[!] Lỗi trong quy trình gửi cho {apt}: {e}")
                        # Quay lại trang chính để reset trạng thái cho lượt sau
                        await page.goto(actual_url)
                        await asyncio.sleep(3)
                        continue

                    except Exception as e:
                        self.log(f"[!] Lỗi khi xử lý Email {apt}: {e}")
                        # Nếu lỗi, quay lại trang chính để chuẩn bị cho lượt sau
                        await page.goto(actual_url)
                        continue

                await browser.close()
        except Exception as e:
            self.log(f"[!] Lỗi hệ thống Email: {e}")
        finally:
            self.is_running = False
            self.log("--- TIẾN TRÌNH EMAIL KẾT THÚC ---")

    def stop(self):
        self.is_running = False
