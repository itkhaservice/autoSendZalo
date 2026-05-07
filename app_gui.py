import os
import sys

# Thêm thư mục node_modules vào path để nhận diện thư viện cục bộ (giống Node.js)
package_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "node_modules")
if os.path.exists(package_path):
    sys.path.insert(0, package_path)

import customtkinter as ctk
import json
import threading
import asyncio
import subprocess
import sys
from tkinter import filedialog, messagebox
from zalo_core import ZaloAutoSender

# Cấu hình giao diện
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Zalo Auto-Sender (v1.0.0) - Nguồn: CMTHANG")
        
        # Khởi động Full màn hình
        self.after(0, lambda: self.state('zoomed'))
        self.geometry("1100x800")

        # Cấu hình Logo/Icon cho Window
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        logo_png = os.path.join(base_path, "Logo512.png")
        logo_ico = os.path.join(base_path, "Logo512.ico")
        
        # 1. Thử nạp .ico cho Taskbar & Window Frame (chuẩn Windows)
        if os.path.exists(logo_ico):
            try:
                self.iconbitmap(logo_ico)
            except Exception as e:
                print(f"Lỗi nạp iconbitmap: {e}")

        # 2. Thử nạp .png làm icon phụ (fallback)
        if os.path.exists(logo_png):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_png)
                icon_img = ImageTk.PhotoImage(img.resize((32, 32)))
                self.wm_iconphoto(True, icon_img)
                self._icon = icon_img 
            except Exception as e:
                print(f"Lỗi nạp logo png: {e}")

        # Cấu hình thư mục lưu trữ dữ liệu cố định trên Disk C
        self.app_data_dir = r"C:\ZaloAutoSenderData"
        if not os.path.exists(self.app_data_dir):
            try:
                os.makedirs(self.app_data_dir)
            except:
                # Nếu không có quyền tạo ở root C, lưu vào AppData của người dùng
                self.app_data_dir = os.path.join(os.environ["APPDATA"], "ZaloAutoSender")
                if not os.path.exists(self.app_data_dir):
                    os.makedirs(self.app_data_dir)

        self.settings_file = os.path.join(self.app_data_dir, "settings.json")
        self.storage_state_file = os.path.join(self.app_data_dir, "storage_state.json")

        # Khởi tạo core với đường dẫn storage_state cố định
        self.zalo_bot = ZaloAutoSender(storage_state=self.storage_state_file, log_callback=self.update_log)
        
        self.settings = self.load_settings()

        self.create_widgets()

        # Đảm bảo Playwright đã được cài đặt driver (chạy ngầm)
        self.ensure_playwright()

    def ensure_playwright(self):
        """Chạy kiểm tra trình duyệt trong luồng riêng để tránh đơ giao diện"""
        def _check():
            try:
                self.update_log("[*] Đang kiểm tra trình duyệt hỗ trợ...")
                
                try:
                    from playwright.cli.main import main as playwright_main
                except ImportError:
                    self.update_log("[!] Cảnh báo: Không thể nạp module cài đặt trình duyệt tự động.")
                    self.update_log("[!] Nếu trình duyệt không mở được, hãy cài đặt Playwright thủ công.")
                    return

                # Lưu lại argv gốc
                original_argv = sys.argv[:]
                
                # Giả lập tham số dòng lệnh cho playwright
                sys.argv = [sys.argv[0], "install", "chromium"]
                
                try:
                    playwright_main()
                    self.update_log("[+] Kiểm tra trình duyệt hoàn tất.")
                except SystemExit:
                    self.update_log("[+] Trình duyệt đã sẵn sàng.")
                except Exception as e:
                    self.update_log(f"[!] Cảnh báo khi cài đặt: {e}")
                finally:
                    # Khôi phục lại argv gốc
                    sys.argv = original_argv
                    
            except Exception as e:
                self.update_log(f"[!] Lỗi kiểm tra trình duyệt: {e}")
        
        threading.Thread(target=_check, daemon=True).start()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "excel_path": "",
            "sheet_name": "04.2026",
            "file_dir": "",
            "img_dir": "",
            "message": "Chào anh/chị căn hộ {apt}, em gửi thông báo phí tháng {month_year} ạ. Đây là tin nhắn tự động nhắc đóng phí đợt {period} (hạn chót đến ngày {deadline}). Anh/chị vui lòng bỏ qua thông báo này nếu đã hoàn tất thanh toán. Em xin cảm ơn!",
            "month_year": "04/2026",
            "period": "1",
            "deadline": "28/04/2026",
            "col_apt": "B",
            "col_sdt": "J",
            "use_attach_file": True,
            "use_attach_img": False
        }

    def save_settings(self):
        self.settings = {
            "excel_path": self.excel_entry.get(),
            "sheet_name": self.sheet_entry.get(),
            "file_dir": self.dir_entry.get(),
            "img_dir": self.img_dir_entry.get(),
            "message": self.message_text.get("1.0", "end-1c"),
            "month_year": self.month_year_entry.get(),
            "period": self.period_entry.get(),
            "deadline": self.deadline_entry.get(),
            "col_apt": self.col_apt_entry.get(),
            "col_sdt": self.col_sdt_entry.get(),
            "use_attach_file": self.use_file_var.get(),
            "use_attach_img": self.use_img_var.get()
        }
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar / Left Panel (Session Management) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="ZALO BOT", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.login_btn = ctk.CTkButton(self.sidebar_frame, text="Đăng nhập (QR)", command=self.handle_login)
        self.login_btn.grid(row=1, column=0, padx=20, pady=10)

        self.check_btn = ctk.CTkButton(self.sidebar_frame, text="Kiểm tra Session", command=self.handle_check_login)
        self.check_btn.grid(row=2, column=0, padx=20, pady=10)

        self.shortcut_btn = ctk.CTkButton(self.sidebar_frame, text="Tạo Shortcut Desktop", fg_color="#34495e", command=self.handle_create_shortcut)
        self.shortcut_btn.grid(row=3, column=0, padx=20, pady=10)

        self.debt_btn = ctk.CTkButton(self.sidebar_frame, text="Tự động gạch nợ", fg_color="#d35400", state="disabled", command=lambda: messagebox.showinfo("Thông báo", "Tính năng này sẽ được cập nhật trong tương lai."))
        self.debt_btn.grid(row=4, column=0, padx=20, pady=10)
        
        self.info_label = ctk.CTkLabel(self.sidebar_frame, text="Lưu ý đặt tên file:\n[Mã căn hộ].pdf\nVí dụ: 01.01.pdf", 
                                       font=ctk.CTkFont(size=12), justify="left", text_color="#3498db")
        self.info_label.grid(row=5, column=0, padx=20, pady=20)

        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Giao diện:", anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=7, column=0, padx=20, pady=(10, 20))
        self.appearance_mode_optionemenu.set("Dark")

        # --- Main Content ---
        self.main_frame = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(1, weight=1)

        # 1. Excel Config
        self.config_label = ctk.CTkLabel(self.main_frame, text="1. Cấu hình dữ liệu nguồn", font=ctk.CTkFont(size=16, weight="bold"))
        self.config_label.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 10), sticky="w")

        ctk.CTkLabel(self.main_frame, text="File Excel:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.excel_entry = ctk.CTkEntry(self.main_frame)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.excel_entry.insert(0, self.settings["excel_path"])
        self.excel_btn = ctk.CTkButton(self.main_frame, text="Chọn", width=60, command=self.browse_excel)
        self.excel_btn.grid(row=1, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.main_frame, text="Tên Sheet:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.sheet_entry = ctk.CTkEntry(self.main_frame)
        self.sheet_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.sheet_entry.insert(0, self.settings["sheet_name"])

        ctk.CTkLabel(self.main_frame, text="Cột Căn hộ:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.col_apt_entry = ctk.CTkEntry(self.main_frame, width=100)
        self.col_apt_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        self.col_apt_entry.insert(0, self.settings.get("col_apt", "B"))

        ctk.CTkLabel(self.main_frame, text="Cột SĐT:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
        self.col_sdt_entry = ctk.CTkEntry(self.main_frame, width=100)
        self.col_sdt_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")
        self.col_sdt_entry.insert(0, self.settings.get("col_sdt", "J"))

        # 2. Files and Images Config
        self.files_label = ctk.CTkLabel(self.main_frame, text="2. Cấu hình Tệp & Hình ảnh", font=ctk.CTkFont(size=16, weight="bold"))
        self.files_label.grid(row=5, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")

        # Thư mục File thông báo
        ctk.CTkLabel(self.main_frame, text="Thư mục Thông báo:").grid(row=6, column=0, padx=10, pady=5, sticky="e")
        self.dir_entry = ctk.CTkEntry(self.main_frame)
        self.dir_entry.grid(row=6, column=1, padx=10, pady=5, sticky="ew")
        self.dir_entry.insert(0, self.settings["file_dir"])
        self.dir_btn = ctk.CTkButton(self.main_frame, text="Chọn", width=60, command=self.browse_dir)
        self.dir_btn.grid(row=6, column=2, padx=10, pady=5)

        self.use_file_var = ctk.BooleanVar(value=self.settings.get("use_attach_file", True))
        self.file_switch = ctk.CTkSwitch(self.main_frame, text="Bật gửi file thông báo (.pdf/.png)", variable=self.use_file_var)
        self.file_switch.grid(row=7, column=1, padx=10, pady=5, sticky="w")

        # Thư mục Hình ảnh bổ sung
        ctk.CTkLabel(self.main_frame, text="Thư mục Hình ảnh:").grid(row=8, column=0, padx=10, pady=5, sticky="e")
        self.img_dir_entry = ctk.CTkEntry(self.main_frame)
        self.img_dir_entry.grid(row=8, column=1, padx=10, pady=5, sticky="ew")
        self.img_dir_entry.insert(0, self.settings.get("img_dir", ""))
        self.img_btn = ctk.CTkButton(self.main_frame, text="Chọn", width=60, command=self.browse_img_dir)
        self.img_btn.grid(row=8, column=2, padx=10, pady=5)

        self.use_img_var = ctk.BooleanVar(value=self.settings.get("use_attach_img", False))
        self.img_switch = ctk.CTkSwitch(self.main_frame, text="Bật gửi hình ảnh bổ sung", variable=self.use_img_var)
        self.img_switch.grid(row=9, column=1, padx=10, pady=5, sticky="w")

        # 3. Dynamic Info Config
        self.dynamic_label = ctk.CTkLabel(self.main_frame, text="3. Thông tin nội dung", font=ctk.CTkFont(size=16, weight="bold"))
        self.dynamic_label.grid(row=10, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")

        ctk.CTkLabel(self.main_frame, text="Tháng/Năm:").grid(row=11, column=0, padx=10, pady=5, sticky="e")
        self.month_year_entry = ctk.CTkEntry(self.main_frame)
        self.month_year_entry.grid(row=11, column=1, padx=10, pady=5, sticky="ew")
        self.month_year_entry.insert(0, self.settings.get("month_year", "04/2026"))

        ctk.CTkLabel(self.main_frame, text="Đợt báo phí:").grid(row=12, column=0, padx=10, pady=5, sticky="e")
        self.period_entry = ctk.CTkEntry(self.main_frame)
        self.period_entry.grid(row=12, column=1, padx=10, pady=5, sticky="ew")
        self.period_entry.insert(0, self.settings.get("period", "1"))

        ctk.CTkLabel(self.main_frame, text="Ngày hạn chót:").grid(row=13, column=0, padx=10, pady=5, sticky="e")
        self.deadline_entry = ctk.CTkEntry(self.main_frame)
        self.deadline_entry.grid(row=13, column=1, padx=10, pady=5, sticky="ew")
        self.deadline_entry.insert(0, self.settings.get("deadline", "28/04/2026"))

        # 4. Message Config
        ctk.CTkLabel(self.main_frame, text="Nội dung tin nhắn:").grid(row=14, column=0, padx=10, pady=10, sticky="ne")
        self.message_text = ctk.CTkTextbox(self.main_frame, height=120)
        self.message_text.grid(row=14, column=1, columnspan=2, padx=10, pady=10, sticky="ew")
        self.message_text.insert("1.0", self.settings["message"])
        ctk.CTkLabel(self.main_frame, text="Từ khóa: {apt}, {month_year}, {period}, {deadline}", font=ctk.CTkFont(size=11, slant="italic")).grid(row=15, column=1, sticky="w", padx=10)

        # 5. Control Buttons
        self.control_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.control_frame.grid(row=16, column=0, columnspan=3, pady=20)
        
        self.start_btn = ctk.CTkButton(self.control_frame, text="BẮT ĐẦU GỬI", font=ctk.CTkFont(size=14, weight="bold"), 
                                       height=40, width=200, command=self.handle_start)
        self.start_btn.pack(side="left", padx=10)

        self.stop_btn = ctk.CTkButton(self.control_frame, text="DỪNG", font=ctk.CTkFont(size=14, weight="bold"), 
                                      height=40, width=100, fg_color="#d35400", hover_color="#e67e22", command=self.handle_stop)
        self.stop_btn.pack(side="left", padx=10)

        # 6. Logs
        self.log_label = ctk.CTkLabel(self.main_frame, text="Nhật ký hoạt động", font=ctk.CTkFont(size=16, weight="bold"))
        self.log_label.grid(row=17, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")
        
        self.log_box = ctk.CTkTextbox(self.main_frame, height=250)
        self.log_box.grid(row=18, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        self.log_box.configure(state="disabled")

    # --- Event Handlers ---

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, filename)

    def browse_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.dir_entry.delete(0, "end")
            self.dir_entry.insert(0, directory)

    def browse_img_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.img_dir_entry.delete(0, "end")
            self.img_dir_entry.insert(0, directory)

    def update_log(self, message):
        # Sử dụng after để đảm bảo cập nhật UI từ main thread
        self.after(0, self._safe_update_log, message)

    def _safe_update_log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def run_async(self, coro):
        def _run():
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(coro)
            loop.close()
        threading.Thread(target=_run, daemon=True).start()

    def handle_login(self):
        self.update_log("[>] Đang khởi tạo tiến trình đăng nhập...")
        self.login_btn.configure(state="disabled")
        
        async def task():
            try:
                await self.zalo_bot.login()
            finally:
                self.after(0, lambda: self.login_btn.configure(state="normal"))
        
        self.run_async(task())

    def handle_check_login(self):
        self.update_log("[>] Đang kiểm tra trạng thái...")
        self.check_btn.configure(state="disabled")
        
        async def task():
            try:
                await self.zalo_bot.check_login()
            finally:
                self.after(0, lambda: self.check_btn.configure(state="normal"))
        
        self.run_async(task())

    def handle_start(self):
        self.save_settings()
        excel = self.excel_entry.get()
        sheet = self.sheet_entry.get()
        fdir = self.dir_entry.get()
        idir = self.img_dir_entry.get()
        msg = self.message_text.get("1.0", "end-1c")
        
        # Lấy thêm các tham số mới
        config = {
            "month_year": self.month_year_entry.get(),
            "period": self.period_entry.get(),
            "deadline": self.deadline_entry.get(),
            "col_apt": self.col_apt_entry.get().strip().upper(),
            "col_sdt": self.col_sdt_entry.get().strip().upper(),
            "use_attach_file": self.use_file_var.get(),
            "use_attach_img": self.use_img_var.get()
        }

        if not excel or not msg:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ cấu hình!")
            return
            
        if config["use_attach_file"] and not fdir:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục chứa file thông báo!")
            return
        
        if config["use_attach_img"] and not idir:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn thư mục chứa hình ảnh bổ sung!")
            return

        self.start_btn.configure(state="disabled")
        
        async def task():
            await self.zalo_bot.send_process(excel, sheet, fdir, idir, msg, config)
            self.after(0, lambda: self.start_btn.configure(state="normal"))

        self.run_async(task())

    def handle_stop(self):
        self.zalo_bot.stop()
        self.update_log("[!] Đang yêu cầu dừng...")

    def handle_create_shortcut(self):
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
            shortcut_path = os.path.join(desktop, "Zalo Auto-Sender.lnk")
            
            # Nếu đang chạy từ bản build
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
                work_dir = os.path.dirname(exe_path)
            else:
                messagebox.showwarning("Cảnh báo", "Bạn đang chạy script Python. Shortcut chỉ nên tạo từ file .exe sau khi đóng gói.")
                return

            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = exe_path
            shortcut.WorkingDirectory = work_dir
            shortcut.IconLocation = exe_path
            shortcut.save()

            messagebox.showinfo("Thành công", f"Đã tạo shortcut ngoài Desktop:\n{shortcut_path}")
            self.update_log("[+] Đã tạo shortcut ngoài Desktop.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo shortcut: {e}")
            self.update_log(f"[!] Lỗi tạo shortcut: {e}")

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()
