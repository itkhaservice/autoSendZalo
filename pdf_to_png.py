import os
from pdf2image import convert_from_path

# CẤU HÌNH ĐƯỜNG DẪN TỰ ĐỘNG
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(BASE_DIR, "NHAC NO THANG 04", "NHAC NO THANG 04")
OUTPUT_DIR = os.path.join(BASE_DIR, "NHAC NO PNG")
# Đường dẫn Poppler nằm ngay trong project
POPPLER_PATH = os.path.join(BASE_DIR, "poppler", "Library", "bin")

def convert_all_pdfs():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"[*] Đã tạo thư mục: {OUTPUT_DIR}")

    if not os.path.exists(PDF_DIR):
        print(f"[!] Lỗi: Không tìm thấy thư mục PDF tại {PDF_DIR}")
        return

    pdf_files = [f for f in os.listdir(PDF_DIR) if f.lower().endswith('.pdf')]
    print(f"[*] Tìm thấy {len(pdf_files)} file PDF cần chuyển đổi.")

    for i, filename in enumerate(pdf_files):
        pdf_path = os.path.join(PDF_DIR, filename)
        name_without_ext = os.path.splitext(filename)[0]
        output_path = os.path.join(OUTPUT_DIR, f"{name_without_ext}.png")

        if os.path.exists(output_path):
            print(f"[-] [{i+1}/{len(pdf_files)}] Đã tồn tại: {name_without_ext}.png. Bỏ qua.")
            continue

        print(f"[>] [{i+1}/{len(pdf_files)}] Đang chuyển đổi: {filename}...")

        try:
            # Chuyển đổi PDF sang hình ảnh (lấy trang 1)
            images = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
            
            if images:
                images[0].save(output_path, 'PNG')
                print(f"[+] Thành công: {filename} -> {name_without_ext}.png")
        except Exception as e:
            print(f"[!] Lỗi khi chuyển đổi {filename}: {e}")

    print("\n--- HOÀN THÀNH ---")

if __name__ == "__main__":
    convert_all_pdfs()
