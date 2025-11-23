import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader

# ================== HÀM ĐẾM PDF =====================
def count_pdf_pages_realtime(folder, update_callback, done_callback):
    total_files = 0
    total_pages = 0

    for root, dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf"):
                total_files += 1
                pdf_path = os.path.join(root, f)

                pages = 0
                try:
                    reader = PdfReader(pdf_path)
                    pages = len(reader.pages)
                    total_pages += pages
                except:
                    pass

                # Cập nhật giao diện mỗi file
                update_callback(total_files, total_pages)

    done_callback(total_files, total_pages)


# ================== HÀM XUẤT TXT =====================
def export_to_txt(folder, total_files, total_pages):
    save_path = filedialog.asksaveasfilename(
        initialfile="thong_ke_pdf.txt",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if not save_path:
        return

    with open(save_path, "w", encoding="utf-8") as f:
        f.write(f"Thư mục: {folder}\n")
        f.write(f"Tổng số file PDF: {total_files}\n")
        f.write(f"Tổng số trang PDF: {total_pages}\n")

    messagebox.showinfo("Hoàn tất", f"Đã lưu: {save_path}")


# ================== CHỌN THƯ MỤC =====================
def select_folder():
    folder = filedialog.askdirectory()
    if folder:
        lbl_folder.config(text=folder)
        btn_start.config(state="normal")  # bật nút đếm


# ================== NÚT BẮT ĐẦU ĐẾM =====================
def start_counting():
    folder = lbl_folder.cget("text")
    if folder == "(Chưa chọn thư mục)":
        messagebox.showwarning("Thông báo", "Bạn chưa chọn thư mục")
        return

    btn_start.config(state="disabled")

    # Tạo luồng (thread) để không treo GUI
    t = threading.Thread(
        target=count_pdf_pages_realtime,
        args=(folder, update_realtime, count_done)
    )
    t.daemon = True
    t.start()


# ================== CALLBACK CẬP NHẬT REALTIME =====================
def update_realtime(file_count, page_count):
    lbl_files.config(text=f"Tổng số file PDF: {file_count}")
    lbl_pages.config(text=f"Tổng số trang PDF: {page_count}")


# ================== CALLBACK KHI HOÀN THÀNH =====================
def count_done(file_count, page_count):
    btn_start.config(state="normal")
    messagebox.showinfo("Xong", "Đã quét xong toàn bộ thư mục!")

    export_to_txt(
        lbl_folder.cget("text"), 
        file_count, 
        page_count
    )


# ================== GIAO DIỆN CHÍNH =====================
window = tk.Tk()
window.title("Thống kê PDF (đa luồng, realtime)")
window.geometry("600x300")

tk.Label(window, text="Chọn thư mục chứa PDF để thống kê", font=("Arial", 14)).pack(pady=10)

btn_folder = tk.Button(window, text="Chọn thư mục", font=("Arial", 12), command=select_folder)
btn_folder.pack(pady=5)

lbl_folder = tk.Label(window, text="(Chưa chọn thư mục)", fg="blue", font=("Arial", 11))
lbl_folder.pack(pady=5)

btn_start = tk.Button(window, text="Bắt đầu đếm", font=("Arial", 12), state="disabled", command=start_counting)
btn_start.pack(pady=10)

lbl_files = tk.Label(window, text="Tổng số file PDF: 0", font=("Arial", 12))
lbl_files.pack()

lbl_pages = tk.Label(window, text="Tổng số trang PDF: 0", font=("Arial", 12))
lbl_pages.pack()

window.mainloop()
