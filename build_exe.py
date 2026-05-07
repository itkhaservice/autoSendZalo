import os
import subprocess
import sys
import customtkinter

def build():
    # Tên ứng dụng
    app_name = "ZaloAutoSender"
    main_script = "app_gui.py"

    if not os.path.exists(main_script):
        print(f"[!] Lỗi: Không tìm thấy file {main_script}")
        return

    # Lấy đường dẫn customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    # Lấy đường dẫn playwright
    import playwright
    pw_path = os.path.dirname(playwright.__file__)
    
    # Lệnh build - Chuyển sang --onedir để chạy nhanh và ổn định hơn
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",
        "--noconsole",
        f"--name={app_name}",
        f"--icon=Logo512.ico",
        f"--add-data={ctk_path};customtkinter/",
        f"--add-data={pw_path};playwright/",
        "--add-data=ms-playwright;ms-playwright/",
        "--add-data=poppler;poppler/",
        "--add-data=Logo512.png;.",
        "--add-data=Logo512.ico;.",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--clean",
        "--noconfirm",
        "--exclude-module=matplotlib",
        "--exclude-module=notebook",
        "--exclude-module=scipy",
        "--exclude-module=torch",
        "--exclude-module=torchvision",
        main_script
    ]

    print(f"[*] Đang khởi chạy PyInstaller...")
    print(f"[*] Bao gồm dữ liệu CustomTkinter từ: {ctk_path}")
    print(f"[*] Đang đóng gói {app_name} (One-Dir)...")
    
    try:
        subprocess.run(cmd, check=True)
        print(f"\n[+] Đóng gói hoàn tất!")
        print(f"[+] Ứng dụng của bạn nằm tại thư mục: {os.path.abspath(os.path.join('dist', app_name))}")
        print(f"[+] File chạy chính: {os.path.abspath(os.path.join('dist', app_name, app_name + '.exe'))}")
    except subprocess.CalledProcessError as e:
        print(f"\n[!] Lỗi trong quá trình build: {e}")

if __name__ == "__main__":
    build()
