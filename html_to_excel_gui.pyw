import pandas as pd
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

class HTMLToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HTML ➜ Excel Converter (Nâng cao)")
        self.root.geometry("900x700")
        self.root.configure(bg="#f4f4f4")

        self.html_path = tk.StringVar()
        self.excel_path = tk.StringVar(value="Output.xlsx")
        self.tables = []
        self.table_vars = []
        self.preview_text = None

        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use("clam")

        self.frame = ttk.Frame(self.root, padding=10)
        self.frame.pack(fill='both', expand=True)

        ttk.Label(self.frame, text="Kéo file HTML vào hoặc chọn thủ công:").grid(row=0, column=0, columnspan=3, sticky='w')

        entry_html = ttk.Entry(self.frame, textvariable=self.html_path, width=80)
        entry_html.grid(row=1, column=0, columnspan=2, pady=5)
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

        ttk.Button(self.frame, text="📂 Duyệt HTML", command=self.choose_html_file).grid(row=1, column=2)

        ttk.Button(self.frame, text="🔁 Tải bảng HTML", command=self.load_tables).grid(row=2, column=0, pady=10)

        self.check_frame = ttk.LabelFrame(self.frame, text="Chọn bảng muốn chuyển:", padding=10)
        self.check_frame.grid(row=3, column=0, columnspan=3, sticky='ew')

        ttk.Button(self.frame, text="✅ Chọn tất cả", command=self.select_all).grid(row=4, column=0, pady=5)
        ttk.Button(self.frame, text="❌ Bỏ chọn tất cả", command=self.deselect_all).grid(row=4, column=1, pady=5)

        ttk.Button(self.frame, text="💾 Chọn nơi lưu Excel", command=self.choose_excel_path).grid(row=5, column=0, columnspan=3)
        ttk.Entry(self.frame, textvariable=self.excel_path, width=80).grid(row=6, column=0, columnspan=3, pady=5)

        ttk.Button(self.frame, text="🚀 Chuyển đổi", command=self.convert).grid(row=7, column=0, columnspan=3, pady=10)

        self.status_label = ttk.Label(self.frame, text="", foreground="green")
        self.status_label.grid(row=8, column=0, columnspan=3)

        ttk.Label(self.frame, text="🔍 Xem trước bảng đã chọn:").grid(row=9, column=0, columnspan=3, sticky='w', pady=(10, 0))
        self.preview_text = tk.Text(self.frame, height=12, wrap='none', font=("Courier", 10))
        self.preview_text.grid(row=10, column=0, columnspan=3, sticky="nsew", pady=5)
        self.frame.rowconfigure(10, weight=1)

    def handle_drop(self, event):
        file_path = event.data.strip('{}')
        if file_path.lower().endswith(('.html', '.htm')):
            self.html_path.set(file_path)
            self.status_label.config(text="Đã nhận file kéo thả!", foreground="blue")
        else:
            messagebox.showwarning("Tệp không hợp lệ", "Vui lòng thả file HTML hợp lệ.")

    def choose_html_file(self):
        path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html;*.htm")])
        if path:
            self.html_path.set(path)

    def choose_excel_path(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path.set(path)

    def load_tables(self):
        path = self.html_path.get()
        self.tables = []
        self.table_vars = []
        for widget in self.check_frame.winfo_children():
            widget.destroy()

        try:
            with open(path, 'r', encoding='utf-8') as f:
                soup = BeautifulSoup(f.read(), 'lxml')

            # Xử lý nâng cao: bỏ bảng rỗng hoặc < 3 dòng
            raw_tables = pd.read_html(str(soup))
            self.tables = [df.dropna(how='all').dropna(axis=1, how='all') for df in raw_tables if df.shape[0] >= 3]

            for idx, table in enumerate(self.tables):
                # Làm sạch tên cột
                table.columns = [str(c).strip() for c in table.columns]

                var = tk.BooleanVar(value=False)
                self.table_vars.append(var)

                def preview_table(index=idx):
                    self.preview_text.delete(1.0, tk.END)
                    preview = self.tables[index].head(10).to_string(index=False)
                    self.preview_text.insert(tk.END, preview)

                chk = ttk.Checkbutton(self.check_frame, text=f"Bảng {idx+1} - {table.shape[0]} dòng, {table.shape[1]} cột", variable=var, command=preview_table)
                chk.pack(anchor='w')

            self.status_label.config(text=f"🔍 Đã phát hiện {len(self.tables)} bảng hợp lệ.", foreground="green")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc HTML: {e}")

    def select_all(self):
        for var in self.table_vars:
            var.set(True)

    def deselect_all(self):
        for var in self.table_vars:
            var.set(False)

    def convert(self):
        selected_tables = [table for i, table in enumerate(self.tables) if self.table_vars[i].get()]
        if not selected_tables:
            messagebox.showwarning("Chưa chọn bảng", "Vui lòng chọn ít nhất 1 bảng để chuyển đổi.")
            return

        excel_file = self.excel_path.get()
        if not excel_file:
            messagebox.showwarning("Chưa chọn nơi lưu", "Vui lòng chọn nơi lưu file Excel.")
            return

        try:
            combined = pd.concat(selected_tables, ignore_index=True)
            combined.to_excel(excel_file, index=False)
            self.status_label.config(text=f"✅ Đã xuất file: {excel_file}", foreground="green")
        except Exception as e:
            self.status_label.config(text=f"❌ Lỗi: {e}", foreground="red")

if __name__ == "__main__":
    app = HTMLToExcelApp(TkinterDnD.Tk())
    app.root.mainloop()
