import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
import os
import math

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Công cụ tách file Excel")
        self.root.geometry("550x350")

        self.file_path = tk.StringVar()
        self.num_parts = tk.StringVar(value="2")
        self.is_processing = False

        # --- Giao diện ---
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill="both", expand=True)

        # Chọn file
        file_frame = ttk.LabelFrame(main_frame, text="Chọn file", padding="10")
        file_frame.pack(fill="x", expand=True)

        ttk.Label(file_frame, text="Đường dẫn file Excel:").pack(anchor="w")
        
        entry_frame = ttk.Frame(file_frame)
        entry_frame.pack(fill="x", expand=True, pady=5)
        
        file_entry = ttk.Entry(entry_frame, textvariable=self.file_path, state="readonly", width=50)
        file_entry.pack(side="left", fill="x", expand=True)
        
        browse_button = ttk.Button(entry_frame, text="...", command=self.browse_file, width=4)
        browse_button.pack(side="left", padx=(5, 0))

        # Cấu hình tách file
        split_frame = ttk.LabelFrame(main_frame, text="Cấu hình", padding="10")
        split_frame.pack(fill="x", expand=True, pady=10)

        ttk.Label(split_frame, text="Số phần muốn chia:").pack(side="left", padx=(0, 10))
        num_parts_entry = ttk.Entry(split_frame, textvariable=self.num_parts, width=10)
        num_parts_entry.pack(side="left")

        # Nút thực thi và thanh trạng thái
        self.run_button = ttk.Button(main_frame, text="Bắt đầu tách file", command=self.start_splitting)
        self.run_button.pack(pady=10, fill="x")

        self.status_label = ttk.Label(main_frame, text="Sẵn sàng", relief="sunken", anchor="w", padding=5)
        self.status_label.pack(fill="x", expand=True, side="bottom")

    def browse_file(self):
        """Mở hộp thoại để chọn file Excel."""
        path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )
        if path:
            self.file_path.set(path)
            self.update_status(f"Đã chọn file: {os.path.basename(path)}")

    def update_status(self, message):
        """Cập nhật thanh trạng thái."""
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def start_splitting(self):
        """Bắt đầu quá trình tách file trong một thread riêng để không làm treo GUI."""
        if self.is_processing:
            return

        file_path = self.file_path.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Lỗi", "Vui lòng chọn một file Excel hợp lệ.")
            return

        try:
            num_parts = int(self.num_parts.get())
            if num_parts <= 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Lỗi", "Số phần muốn chia phải là một số nguyên lớn hơn 1.")
            return

        self.is_processing = True
        self.run_button.config(state="disabled", text="Đang xử lý...")
        
        # Chạy tác vụ trong thread khác
        threading.Thread(
            target=self.split_excel_task,
            args=(file_path, num_parts),
            daemon=True
        ).start()

    def split_excel_task(self, file_path, num_parts):
        """Hàm logic để tách file Excel."""
        try:
            self.update_status("Đang đọc file Excel...")
            df = pd.read_excel(file_path, engine='openpyxl')
            
            if df.empty:
                raise ValueError("File Excel rỗng.")

            num_rows = len(df)
            if num_rows < num_parts:
                messagebox.showwarning("Cảnh báo", f"Số dòng ({num_rows}) ít hơn số phần muốn chia ({num_parts}). Không thể chia file.")
                self.reset_ui()
                return
            
            # Tính toán số dòng mỗi file
            chunk_size = math.ceil(num_rows / num_parts)
            
            file_dir, file_name = os.path.split(file_path)
            base_name, ext = os.path.splitext(file_name)

            self.update_status(f"Tổng số {num_rows} dòng, chia thành {num_parts} phần...")

            for i in range(num_parts):
                start_row = i * chunk_size
                end_row = start_row + chunk_size
                chunk_df = df.iloc[start_row:end_row]

                if chunk_df.empty:
                    continue

                output_filename = os.path.join(file_dir, f"{base_name}_part_{i + 1}{ext}")
                
                self.update_status(f"Đang tạo file phần {i + 1}/{num_parts}...")
                
                # Ghi ra file mới, giữ nguyên header
                chunk_df.to_excel(output_filename, index=False, engine='openpyxl')

            self.update_status("Hoàn tất! Các file đã được lưu cùng thư mục với file gốc.")
            messagebox.showinfo("Thành công", f"Đã tách file thành công thành {num_parts} phần.")

        except ImportError:
             messagebox.showerror("Lỗi Thiếu Thư Viện", 
                                  "Không tìm thấy thư viện 'pandas' hoặc 'openpyxl'.\n"
                                  "Vui lòng cài đặt bằng cách chạy lệnh:\n"
                                  "pip install pandas openpyxl")
        except Exception as e:
            messagebox.showerror("Đã có lỗi xảy ra", f"Chi tiết lỗi: {e}")
            self.update_status("Sẵn sàng")
        finally:
            self.reset_ui()
            
    def reset_ui(self):
        """Reset lại giao diện sau khi xử lý xong."""
        self.is_processing = False
        self.run_button.config(state="normal", text="Bắt đầu tách file")

if __name__ == "__main__":
    try:
        import pandas
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Lỗi Thiếu Thư Viện", 
                             "Không tìm thấy thư viện 'pandas' hoặc 'openpyxl'.\n"
                             "Vui lòng cài đặt bằng cách chạy lệnh:\n"
                             "pip install pandas openpyxl")
        root.destroy()
    else:
        root = tk.Tk()
        app = ExcelSplitterApp(root)
        root.mainloop()
