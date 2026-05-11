import os
import subprocess
import sys
import customtkinter

def build():
    # Tên ứng dụng
    app_name = "ZaloAutoSender"
    main_script = "app_gui.py"
    base_dir = os.path.dirname(os.path.abspath(__file__))

    if not os.path.exists(os.path.join(base_dir, main_script)):
        print(f"[!] Lỗi: Không tìm thấy file {main_script}")
        return

    # Lấy đường dẫn customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    
    # Lấy đường dẫn playwright
    import playwright
    pw_path = os.path.dirname(playwright.__file__)
    
    # Lệnh build - Chuyển sang --onedir để chạy nhanh và ổn định hơn
    # Dùng dấu ; cho Windows, : cho Linux/Mac
    sep = ";"
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",
        "--noconsole",
        f"--name={app_name}",
        f"--icon={os.path.join(base_dir, 'Logo512.ico')}",
        f"--add-data={ctk_path}{sep}customtkinter/",
        f"--add-data={pw_path}{sep}playwright/",
        f"--add-data={os.path.join(base_dir, 'pw-browsers')}{sep}pw-browsers/",
        f"--add-data={os.path.join(base_dir, 'poppler')}{sep}poppler/",
        f"--add-data={os.path.join(base_dir, 'mailauto')}{sep}mailauto/",
        f"--add-data={os.path.join(base_dir, 'Logo512.png')}{sep}.",
        f"--add-data={os.path.join(base_dir, 'Logo512.ico')}{sep}.",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=email_core",
        "--hidden-import=zalo_core",
        "--clean",
        "--noconfirm",
        "--exclude-module=matplotlib",
        "--exclude-module=notebook",
        "--exclude-module=scipy",
        "--exclude-module=torch",
        "--exclude-module=torchvision",
        main_script
    ]

    print(f"[*] Đang đóng gói {app_name} (One-Dir)...")
    print(f"[*] Command: {' '.join(cmd)}")
    
    try:
        # Chạy trực tiếp để thấy output
        result = subprocess.run(cmd, capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print("[!] STDERR:")
            print(result.stderr)
            
        if result.returncode == 0:
            print(f"\n[+] Đóng gói hoàn tất!")
            print(f"[+] Ứng dụng của bạn nằm tại thư mục: {os.path.abspath(os.path.join('dist', app_name))}")
            print(f"[+] File chạy chính: {os.path.abspath(os.path.join('dist', app_name, app_name + '.exe'))}")
        else:
            print(f"\n[!] Build thất bại với mã thoát: {result.returncode}")
            
    except Exception as e:
        print(f"\n[!] Lỗi ngoại lệ: {e}")

if __name__ == "__main__":
    build()
