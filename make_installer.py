import os
import subprocess
import winreg

def find_inno_setup():
    """Tìm đường dẫn cài đặt của Inno Setup trong Registry"""
    paths = [
        r"Software\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1",
        r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1"
    ]
    for path in paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                install_location = winreg.QueryValueEx(key, "InstallLocation")[0]
                iscc = os.path.join(install_location, "ISCC.exe")
                if os.path.exists(iscc):
                    return iscc
        except:
            continue
    return None

def make_installer():
    iss_file = "installer_setup.iss"
    if not os.path.exists(iss_file):
        print(f"[!] Không tìm thấy file {iss_file}")
        return

    iscc_path = find_inno_setup()
    if not iscc_path:
        # Thử tìm trong các đường dẫn mặc định
        default_paths = [
            r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
            r"C:\Program Files\Inno Setup 6\ISCC.exe"
        ]
        for p in default_paths:
            if os.path.exists(p):
                iscc_path = p
                break
    
    if not iscc_path:
        print("[!] Không tìm thấy Inno Setup (ISCC.exe).")
        print("[*] Vui lòng cài đặt Inno Setup từ: https://jrsoftware.org/isdl.php")
        return

    print(f"[*] Đang biên dịch bộ cài đặt bằng: {iscc_path}")
    try:
        subprocess.run([iscc_path, iss_file], check=True)
        print("\n" + "="*50)
        print("[+] THÀNH CÔNG! Bản cài đặt của bạn đã sẵn sàng.")
        print(f"[+] File Setup: {os.path.abspath(os.path.join('Output', 'ZaloAutoSender_Setup.exe'))}")
        print("[+] Bạn có thể gửi file này cho bất kỳ máy tính nào để cài đặt.")
        print("="*50)
    except Exception as e:
        print(f"[!] Lỗi khi biên dịch: {e}")

if __name__ == "__main__":
    make_installer()
