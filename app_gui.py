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
from tkinter import filedialog, messagebox

# Import core sau khi đã setup sys.path
try:
    from zalo_core import ZaloAutoSender
    from email_core import EmailAutoSender
except ImportError:
    # Trường hợp đang build hoặc lỗi môi trường
    ZaloAutoSender = None
    EmailAutoSender = None

class App(ctk.CTk):
    # ... (Toàn bộ class App giữ nguyên như cũ, chỉ dọn dẹp hàm __init__ một chút)
    def __init__(self):
        super().__init__()
        # Cấu hình giao diện (Nên để trong __init__ của App)
        try:
            ctk.set_appearance_mode("Dark")
            ctk.set_default_color_theme("green")
        except: pass

        self.title("Zalo & Email Auto-Sender (v1.1.2) - Nguồn: CMTHANG")
        # ... rest of init ...
        
        # Khởi động Full màn hình
        self.after(0, lambda: self.state('zoomed'))
        self.geometry("1200x900")

        # Cấu hình Logo/Icon cho Window
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        logo_png = os.path.join(base_path, "Logo512.png")
        logo_ico = os.path.join(base_path, "Logo512.ico")
        
        if os.path.exists(logo_ico):
            try:
                self.iconbitmap(logo_ico)
            except: pass

        if os.path.exists(logo_png):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_png)
                icon_img = ImageTk.PhotoImage(img.resize((32, 32)))
                self.wm_iconphoto(True, icon_img)
                self._icon = icon_img 
            except: pass

        # Cấu hình thư mục lưu trữ dữ liệu
        self.app_data_dir = r"C:\ZaloAutoSenderData"
        if not os.path.exists(self.app_data_dir):
            try:
                os.makedirs(self.app_data_dir)
            except:
                self.app_data_dir = os.path.join(os.environ["APPDATA"], "ZaloAutoSender")
                if not os.path.exists(self.app_data_dir):
                    os.makedirs(self.app_data_dir)

        self.settings_file = os.path.join(self.app_data_dir, "settings_v2.json")
        self.zalo_storage = os.path.join(self.app_data_dir, "storage_state.json")
        self.email_storage = os.path.join(self.app_data_dir, "email_storage_state.json")

        # Khởi tạo core
        self.zalo_bot = ZaloAutoSender(storage_state=self.zalo_storage, log_callback=self.update_log)
        self.email_bot = EmailAutoSender(storage_state=self.email_storage, log_callback=self.update_log)
        
        self.settings = self.load_settings()

        self.create_widgets()
        self.ensure_playwright()

    def load_settings(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "zalo": {
                "excel_path": "",
                "sheet_name": "Sheet1",
                "file_dir": "",
                "img_dir": "",
                "message": "Chào anh/chị căn hộ {apt}, em gửi thông báo phí tháng {month_year} ạ. Đây là tin nhắn tự động nhắc đóng phí đợt {period} (hạn chót đến ngày {deadline}). Anh/chị vui lòng bỏ qua thông báo này nếu đã hoàn tất thanh toán. Em xin cảm ơn!",
                "month_year": "04/2026",
                "period": "1",
                "deadline": "28/04/2026",
                "col_apt": "B", "col_sdt": "J",
                "use_attach_file": True, "use_attach_img": False
            },
            "email": {
                "url": "https://pro55.emailserver.vn:2096/",
                "user": "bqlrivapark@khaservice.com.vn",
                "pass": "Khas@123",
                "excel_path": "",
                "sheet_name": "Sheet1",
                "file_dir": "",
                "subject": "[THÔNG BÁO PHÍ] Căn hộ {apt} - Tháng {month_year}",
                "body": "Chào anh/chị căn hộ {apt},\n\nBan Quản lý xin gửi thông báo phí tháng {month_year} đến anh/chị. Đây là đợt báo phí thứ {period}, hạn chót thanh toán là ngày {deadline}.\n\nAnh/chị vui lòng kiểm tra file đính kèm và hoàn tất thanh toán đúng hạn. Nếu đã thanh toán, vui lòng bỏ qua thông báo này.\n\nTrân trọng!",
                "month_year": "04/2026",
                "period": "1",
                "deadline": "28/04/2026",
                "col_apt": "B", "col_email": "K"
            }
        }

    def save_settings(self):
        # Thu thập dữ liệu từ UI trước khi lưu
        self.settings["zalo"] = {
            "excel_path": self.z_excel_entry.get(),
            "sheet_name": self.z_sheet_entry.get(),
            "file_dir": self.z_dir_entry.get(),
            "img_dir": self.z_img_dir_entry.get(),
            "message": self.z_message_text.get("1.0", "end-1c"),
            "month_year": self.z_month_year_entry.get(),
            "period": self.z_period_entry.get(),
            "deadline": self.z_deadline_entry.get(),
            "col_apt": self.z_col_apt_entry.get().upper(),
            "col_sdt": self.z_col_sdt_entry.get().upper(),
            "use_attach_file": self.z_use_file_var.get(),
            "use_attach_img": self.z_use_img_var.get()
        }
        self.settings["email"] = {
            "url": self.e_url_entry.get(),
            "user": self.e_user_entry.get(),
            "pass": self.e_pass_entry.get(),
            "excel_path": self.e_excel_entry.get(),
            "sheet_name": self.e_sheet_entry.get(),
            "file_dir": self.e_dir_entry.get(),
            "subject": self.e_subject_entry.get(),
            "body": self.e_body_text.get("1.0", "end-1c"),
            "month_year": self.e_month_year_entry.get(),
            "period": self.e_period_entry.get(),
            "deadline": self.e_deadline_entry.get(),
            "col_apt": self.e_col_apt_entry.get().upper(),
            "col_email": self.e_col_email_entry.get().upper()
        }
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar ---
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="MULTI BOT", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.shortcut_btn = ctk.CTkButton(self.sidebar_frame, text="Tạo Shortcut Desktop", fg_color="#34495e", command=self.handle_create_shortcut)
        self.shortcut_btn.grid(row=1, column=0, padx=20, pady=20)

        # --- Main View (Tabs) ---
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        
        self.tab_zalo = self.tabview.add("GỬI ZALO")
        self.tab_email = self.tabview.add("GỬI EMAIL")
        
        self.setup_zalo_tab()
        self.setup_email_tab()

        # --- Logs Area (Shared) ---
        self.log_frame = ctk.CTkFrame(self, height=200)
        self.log_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self.log_frame, text="Nhật ký hoạt động hệ thống", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, sticky="w")
        self.log_box = ctk.CTkTextbox(self.log_frame, height=150)
        self.log_box.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.log_box.configure(state="disabled")

    def setup_zalo_tab(self):
        self.tab_zalo.grid_columnconfigure(0, weight=1)
        self.tab_zalo.grid_rowconfigure(0, weight=1)
        
        # Sử dụng ScrollableFrame bên trong Tab để giữ style cũ
        container = ctk.CTkScrollableFrame(self.tab_zalo, corner_radius=0, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(1, weight=1)

        # 1. Zalo Session
        self.z_login_btn = ctk.CTkButton(container, text="Đăng nhập Zalo Web (QR)", command=self.handle_zalo_login)
        self.z_login_btn.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

        # 2. Excel Config
        ctk.CTkLabel(container, text="1. Cấu hình dữ liệu nguồn (Zalo)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=1, column=0, columnspan=3, padx=10, pady=(10, 10), sticky="w")

        ctk.CTkLabel(container, text="File Excel:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.z_excel_entry = ctk.CTkEntry(container)
        self.z_excel_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.z_excel_entry.insert(0, self.settings["zalo"]["excel_path"])
        ctk.CTkButton(container, text="Chọn", width=60, command=lambda: self.browse_excel(self.z_excel_entry)).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(container, text="Tên Sheet:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.z_sheet_entry = ctk.CTkEntry(container)
        self.z_sheet_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.z_sheet_entry.insert(0, self.settings["zalo"]["sheet_name"])

        ctk.CTkLabel(container, text="Cột Căn hộ:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
        self.z_col_apt_entry = ctk.CTkEntry(container, width=100)
        self.z_col_apt_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")
        self.z_col_apt_entry.insert(0, self.settings["zalo"]["col_apt"])

        ctk.CTkLabel(container, text="Cột SĐT:").grid(row=5, column=0, padx=10, pady=5, sticky="e")
        self.z_col_sdt_entry = ctk.CTkEntry(container, width=100)
        self.z_col_sdt_entry.grid(row=5, column=1, padx=10, pady=5, sticky="w")
        self.z_col_sdt_entry.insert(0, self.settings["zalo"]["col_sdt"])

        # 3. Files and Images Config
        ctk.CTkLabel(container, text="2. Cấu hình Tệp & Hình ảnh", font=ctk.CTkFont(size=16, weight="bold")).grid(row=6, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")

        ctk.CTkLabel(container, text="Thư mục Thông báo:").grid(row=7, column=0, padx=10, pady=5, sticky="e")
        self.z_dir_entry = ctk.CTkEntry(container)
        self.z_dir_entry.grid(row=7, column=1, padx=10, pady=5, sticky="ew")
        self.z_dir_entry.insert(0, self.settings["zalo"]["file_dir"])
        ctk.CTkButton(container, text="Chọn", width=60, command=lambda: self.browse_dir(self.z_dir_entry)).grid(row=7, column=2, padx=10, pady=5)

        self.z_use_file_var = ctk.BooleanVar(value=self.settings["zalo"]["use_attach_file"])
        ctk.CTkSwitch(container, text="Bật gửi file thông báo (.pdf/.png)", variable=self.z_use_file_var).grid(row=8, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(container, text="Thư mục Hình ảnh:").grid(row=9, column=0, padx=10, pady=5, sticky="e")
        self.z_img_dir_entry = ctk.CTkEntry(container)
        self.z_img_dir_entry.grid(row=9, column=1, padx=10, pady=5, sticky="ew")
        self.z_img_dir_entry.insert(0, self.settings["zalo"]["img_dir"])
        ctk.CTkButton(container, text="Chọn", width=60, command=lambda: self.browse_dir(self.z_img_dir_entry)).grid(row=9, column=2, padx=10, pady=5)

        self.z_use_img_var = ctk.BooleanVar(value=self.settings["zalo"]["use_attach_img"])
        ctk.CTkSwitch(container, text="Bật gửi hình ảnh bổ sung", variable=self.z_use_img_var).grid(row=10, column=1, padx=10, pady=5, sticky="w")

        # 3. Dynamic Info Config
        ctk.CTkLabel(container, text="3. Thông tin nội dung (Zalo)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=11, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")
        ctk.CTkLabel(container, text="Tháng/Năm:").grid(row=12, column=0, padx=10, pady=5, sticky="e")
        self.z_month_year_entry = ctk.CTkEntry(container)
        self.z_month_year_entry.grid(row=12, column=1, padx=10, pady=5, sticky="ew")
        self.z_month_year_entry.insert(0, self.settings["zalo"].get("month_year", "04/2026"))
        
        ctk.CTkLabel(container, text="Đợt báo phí:").grid(row=13, column=0, padx=10, pady=5, sticky="e")
        self.z_period_entry = ctk.CTkEntry(container)
        self.z_period_entry.grid(row=13, column=1, padx=10, pady=5, sticky="ew")
        self.z_period_entry.insert(0, self.settings["zalo"].get("period", "1"))
        
        ctk.CTkLabel(container, text="Ngày hạn chót:").grid(row=14, column=0, padx=10, pady=5, sticky="e")
        self.z_deadline_entry = ctk.CTkEntry(container)
        self.z_deadline_entry.grid(row=14, column=1, padx=10, pady=5, sticky="ew")
        self.z_deadline_entry.insert(0, self.settings["zalo"].get("deadline", "28/04/2026"))

        # 4. Message Config
        ctk.CTkLabel(container, text="4. Nội dung tin nhắn:").grid(row=15, column=0, padx=10, pady=10, sticky="ne")
        self.z_message_text = ctk.CTkTextbox(container, height=120)
        self.z_message_text.grid(row=15, column=1, columnspan=2, padx=10, pady=10, sticky="ew")
        self.z_message_text.insert("1.0", self.settings["zalo"]["message"])
        ctk.CTkLabel(container, text="Từ khóa: {apt}, {month_year}, {period}, {deadline}", font=ctk.CTkFont(size=11, slant="italic")).grid(row=16, column=1, sticky="w", padx=10)

        # 5. Control Buttons
        z_ctrl_frame = ctk.CTkFrame(container, fg_color="transparent")
        z_ctrl_frame.grid(row=17, column=0, columnspan=3, pady=20)
        self.z_start_btn = ctk.CTkButton(z_ctrl_frame, text="BẮT ĐẦU GỬI ZALO", font=ctk.CTkFont(size=14, weight="bold"), height=40, width=200, fg_color="green", command=self.handle_zalo_start)
        self.z_start_btn.pack(side="left", padx=10)
        self.z_stop_btn = ctk.CTkButton(z_ctrl_frame, text="DỪNG", font=ctk.CTkFont(size=14, weight="bold"), height=40, width=100, fg_color="#d35400", command=lambda: self.zalo_bot.stop())
        self.z_stop_btn.pack(side="left", padx=10)

    def setup_email_tab(self):
        self.tab_email.grid_columnconfigure(0, weight=1)
        self.tab_email.grid_rowconfigure(0, weight=1)
        
        container = ctk.CTkScrollableFrame(self.tab_email, corner_radius=0, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_columnconfigure(1, weight=1)

        # 1. Webmail Session
        ctk.CTkLabel(container, text="Tài khoản Webmail:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.e_url_entry = ctk.CTkEntry(container)
        self.e_url_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        self.e_url_entry.insert(0, self.settings["email"]["url"])

        self.e_user_entry = ctk.CTkEntry(container, placeholder_text="Email")
        self.e_user_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.e_user_entry.insert(0, self.settings["email"]["user"])

        self.e_pass_entry = ctk.CTkEntry(container, placeholder_text="Mật khẩu", show="*")
        self.e_pass_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.e_pass_entry.insert(0, self.settings["email"]["pass"])

        self.e_login_btn = ctk.CTkButton(container, text="Đăng nhập Webmail", command=self.handle_email_login)
        self.e_login_btn.grid(row=1, column=2, padx=10, pady=5)

        # 2. Excel Config
        ctk.CTkLabel(container, text="1. Cấu hình dữ liệu nguồn (Email)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=2, column=0, columnspan=3, padx=10, pady=(10, 10), sticky="w")

        ctk.CTkLabel(container, text="File Excel:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.e_excel_entry = ctk.CTkEntry(container)
        self.e_excel_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.e_excel_entry.insert(0, self.settings["email"]["excel_path"])
        ctk.CTkButton(container, text="Chọn", width=60, command=lambda: self.browse_excel(self.e_excel_entry)).grid(row=3, column=2, padx=10, pady=5)

        ctk.CTkLabel(container, text="Tên Sheet:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
        self.e_sheet_entry = ctk.CTkEntry(container)
        self.e_sheet_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.e_sheet_entry.insert(0, self.settings["email"]["sheet_name"])

        ctk.CTkLabel(container, text="Cột Căn hộ:").grid(row=5, column=0, padx=10, pady=5, sticky="e")
        self.e_col_apt_entry = ctk.CTkEntry(container, width=100)
        self.e_col_apt_entry.grid(row=5, column=1, padx=10, pady=5, sticky="w")
        self.e_col_apt_entry.insert(0, self.settings["email"]["col_apt"])

        ctk.CTkLabel(container, text="Cột Email:").grid(row=6, column=0, padx=10, pady=5, sticky="e")
        self.e_col_email_entry = ctk.CTkEntry(container, width=100)
        self.e_col_email_entry.grid(row=6, column=1, padx=10, pady=5, sticky="w")
        self.e_col_email_entry.insert(0, self.settings["email"]["col_email"])

        # 3. Files Config
        ctk.CTkLabel(container, text="2. Thư mục Thông báo (PDF/PNG):").grid(row=7, column=0, padx=10, pady=5, sticky="e")
        self.e_dir_entry = ctk.CTkEntry(container)
        self.e_dir_entry.grid(row=7, column=1, padx=10, pady=5, sticky="ew")
        self.e_dir_entry.insert(0, self.settings["email"]["file_dir"])
        ctk.CTkButton(container, text="Chọn", width=60, command=lambda: self.browse_dir(self.e_dir_entry)).grid(row=7, column=2, padx=10, pady=5)

        # 3. Dynamic Info Config
        ctk.CTkLabel(container, text="3. Thông tin nội dung (Email)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=8, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")
        ctk.CTkLabel(container, text="Tháng/Năm:").grid(row=9, column=0, padx=10, pady=5, sticky="e")
        self.e_month_year_entry = ctk.CTkEntry(container)
        self.e_month_year_entry.grid(row=9, column=1, padx=10, pady=5, sticky="ew")
        self.e_month_year_entry.insert(0, self.settings["email"].get("month_year", "04/2026"))
        
        ctk.CTkLabel(container, text="Đợt báo phí:").grid(row=10, column=0, padx=10, pady=5, sticky="e")
        self.e_period_entry = ctk.CTkEntry(container)
        self.e_period_entry.grid(row=10, column=1, padx=10, pady=5, sticky="ew")
        self.e_period_entry.insert(0, self.settings["email"].get("period", "1"))
        
        ctk.CTkLabel(container, text="Ngày hạn chót:").grid(row=11, column=0, padx=10, pady=5, sticky="e")
        self.e_deadline_entry = ctk.CTkEntry(container)
        self.e_deadline_entry.grid(row=11, column=1, padx=10, pady=5, sticky="ew")
        self.e_deadline_entry.insert(0, self.settings["email"].get("deadline", "28/04/2026"))

        # 4. Email Content Config
        ctk.CTkLabel(container, text="4. Nội dung Email:", font=ctk.CTkFont(size=16, weight="bold")).grid(row=12, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="w")
        
        ctk.CTkLabel(container, text="Tiêu đề:").grid(row=13, column=0, padx=10, pady=5, sticky="e")
        self.e_subject_entry = ctk.CTkEntry(container)
        self.e_subject_entry.grid(row=13, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        self.e_subject_entry.insert(0, self.settings["email"]["subject"])

        ctk.CTkLabel(container, text="Nội dung:").grid(row=14, column=0, padx=10, pady=10, sticky="ne")
        self.e_body_text = ctk.CTkTextbox(container, height=120)
        self.e_body_text.grid(row=14, column=1, columnspan=2, padx=10, pady=10, sticky="ew")
        self.e_body_text.insert("1.0", self.settings["email"]["body"])
        ctk.CTkLabel(container, text="Từ khóa: {apt}, {month_year}, {period}, {deadline}", font=ctk.CTkFont(size=11, slant="italic")).grid(row=15, column=1, sticky="w", padx=10)

        # 5. Control Buttons
        e_ctrl_frame = ctk.CTkFrame(container, fg_color="transparent")
        e_ctrl_frame.grid(row=16, column=0, columnspan=3, pady=20)
        self.e_start_btn = ctk.CTkButton(e_ctrl_frame, text="BẮT ĐẦU GỬI EMAIL", font=ctk.CTkFont(size=14, weight="bold"), height=40, width=200, fg_color="blue", command=self.handle_email_start)
        self.e_start_btn.pack(side="left", padx=10)
        self.e_stop_btn = ctk.CTkButton(e_ctrl_frame, text="DỪNG", font=ctk.CTkFont(size=14, weight="bold"), height=40, width=100, fg_color="#d35400", command=lambda: self.email_bot.stop())
        self.e_stop_btn.pack(side="left", padx=10)

    # --- Handlers ---
    def handle_zalo_login(self):
        self.z_login_btn.configure(state="disabled")
        async def task():
            try: await self.zalo_bot.login()
            finally: self.after(0, lambda: self.z_login_btn.configure(state="normal"))
        self.run_async(task())

    def handle_email_login(self):
        url = self.e_url_entry.get()
        user = self.e_user_entry.get()
        pw = self.e_pass_entry.get()
        if not url or not user or not pw:
            messagebox.showwarning("Lỗi", "Vui lòng nhập đầy đủ thông tin Webmail")
            return
        self.e_login_btn.configure(state="disabled")
        async def task():
            try: await self.email_bot.login(url, user, pw)
            finally: self.after(0, lambda: self.e_login_btn.configure(state="normal"))
        self.run_async(task())

    def handle_zalo_start(self):
        self.save_settings()
        cfg = self.settings["zalo"].copy() 
        self.z_start_btn.configure(state="disabled")
        async def task():
            await self.zalo_bot.send_process(self.z_excel_entry.get(), self.z_sheet_entry.get(), 
                                           self.z_dir_entry.get(), self.z_img_dir_entry.get(), 
                                           self.z_message_text.get("1.0", "end-1c"), cfg)
            self.after(0, lambda: self.z_start_btn.configure(state="normal"))
        self.run_async(task())

    def handle_email_start(self):
        self.save_settings()
        cfg = self.settings["email"].copy() 
        self.e_start_btn.configure(state="disabled")
        async def task():
            await self.email_bot.send_process(self.e_excel_entry.get(), self.e_sheet_entry.get(), 
                                            self.e_dir_entry.get(), self.e_url_entry.get(),
                                            self.e_subject_entry.get(), self.e_body_text.get("1.0", "end-1c"), cfg)
            self.after(0, lambda: self.e_start_btn.configure(state="normal"))
        self.run_async(task())

    def browse_excel(self, entry):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if f: entry.delete(0, "end"); entry.insert(0, f)

    def browse_dir(self, entry):
        d = filedialog.askdirectory()
        if d: entry.delete(0, "end"); entry.insert(0, d)

    def ensure_playwright(self):
        def _check():
            try:
                self.update_log("[*] Đang kiểm tra trình duyệt Playwright...")
                import subprocess
                # Sử dụng lệnh python -m playwright install thay vì gọi trực tiếp cli.main
                # Điều này giúp tránh lỗi thiếu module playwright.cli trong một số phiên bản
                result = subprocess.run(
                    [sys.executable, "-m", "playwright", "install", "chromium"],
                    capture_output=True,
                    text=True,
                    env=os.environ.copy()
                )
                if result.returncode == 0:
                    self.update_log("[+] Trình duyệt đã sẵn sàng.")
                else:
                    self.update_log(f"[!] Lỗi cài đặt trình duyệt: {result.stderr}")
            except Exception as e:
                self.update_log(f"[!] Lỗi kiểm tra Playwright: {e}")
        threading.Thread(target=_check, daemon=True).start()

    def update_log(self, message):
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

    def handle_create_shortcut(self):
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            desktop = os.path.join(os.environ["USERPROFILE"], "Desktop")
            shortcut_path = os.path.join(desktop, "Zalo & Email Auto-Sender.lnk")
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
                work_dir = os.path.dirname(exe_path)
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = exe_path
                shortcut.WorkingDirectory = work_dir
                shortcut.IconLocation = exe_path
                shortcut.save()
                messagebox.showinfo("Thành công", "Đã tạo shortcut ngoài Desktop.")
            else:
                messagebox.showwarning("Cảnh báo", "Shortcut chỉ tạo được khi chạy từ file .exe")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo shortcut: {e}")


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    
    # Ghi log debug đối số (giúp chẩn đoán lỗi mở liên tục)
    try:
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"\n[DEBUG] Khởi động với args: {sys.argv}\n")
    except: pass

    # CHẶN NHÂN BẢN: Kiểm tra kỹ các tham số tiến trình con
    is_worker = False
    worker_args = ["run-driver", "-m", "--switched-process", "--type=", "inspect", "reinspect"]
    for arg in sys.argv[1:]:
        for w_arg in worker_args:
            if arg.startswith(w_arg):
                is_worker = True
                break
        if is_worker: break

    # Nếu Playwright hoặc Multiprocessing gọi, không khởi tạo GUI
    if is_worker:
        # Nếu là lệnh python -m playwright, module playwright sẽ tự xử lý qua __main__
        pass
    else:
        try:
            app = App()
            app.mainloop()
        except Exception as e:
            import traceback
            try:
                with open("error_log.txt", "a", encoding="utf-8") as f:
                    f.write("\n" + "="*50 + "\n")
                    f.write(f"Lỗi khởi động App: {e}\n")
                    f.write(traceback.format_exc())
            except: pass
            print(f"Lỗi khởi động: {e}")
