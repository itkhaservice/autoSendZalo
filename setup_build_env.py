import os
import subprocess
import sys

def setup():
    package_dir = "node_modules"
    print(f"[*] Đang tải thư viện trực tiếp vào dự án (thư mục '{package_dir}')...")
    
    if not os.path.exists(package_dir):
        os.makedirs(package_dir)

    # Cài đặt thư viện vào thư mục node_modules
    print("[*] Đang cài đặt các thư viện thiết yếu...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "--target", package_dir], check=True)
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller", "--target", package_dir], check=True)

    # Cài đặt playwright browser vào thư mục cục bộ của dự án
    print("[*] Đang tải trình duyệt Playwright vào dự án...")
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.getcwd(), "pw-browsers")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)

    print("\n[+] HOÀN TẤT!")
    print(f"[!] Thư viện đã nằm sẵn trong '{package_dir}'.")
    print(f"[!] Bạn có thể build ngay bằng lệnh:")
    print(f"    python build_exe.py")

if __name__ == "__main__":
    setup()
