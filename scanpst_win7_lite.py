import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    from pywinauto.application import Application
    from pywinauto.findwindows import ElementNotFoundError
    from pywinauto.timings import TimeoutError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

OFFICE_PATHS = {
    "Office 2010": r"C:\\Program Files\\Microsoft Office\\Office14\\SCANPST.EXE",
    "Office 2013": r"C:\\Program Files\\Microsoft Office\\Office15\\SCANPST.EXE",
    "Office 2016 trở lên": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\SCANPST.EXE",
}

class ScanPstLiteWin7:
    def __init__(self, root):
        self.root = root
        self.root.title("ScanPST Lite (Windows 7)")
        self.root.geometry("650x500")

        self.file_path = tk.StringVar()
        self.office_version = tk.StringVar()

        frame = ttk.Frame(root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Chọn phiên bản Office có SCANPST.EXE:").pack(anchor='w')
        self.office_cb = ttk.Combobox(frame, textvariable=self.office_version, state='readonly')
        self.office_cb.pack(fill='x', pady=5)

        self.office_cb['values'] = [k for k, v in OFFICE_PATHS.items() if os.path.exists(v)]
        if self.office_cb['values']:
            self.office_version.set(self.office_cb['values'][0])

        ttk.Label(frame, text="Chọn file PST hoặc OST:").pack(anchor='w', pady=(10,0))
        path_frame = ttk.Frame(frame)
        path_frame.pack(fill='x')
        self.path_entry = ttk.Entry(path_frame, textvariable=self.file_path)
        self.path_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(path_frame, text="...", command=self.choose_file).pack(side='left', padx=5)

        self.start_btn = ttk.Button(frame, text="Bắt đầu sửa lỗi", command=self.start_repair)
        self.start_btn.pack(pady=10, fill='x')

        self.log_area = scrolledtext.ScrolledText(frame, height=15, state='disabled', font=("Consolas", 9))
        self.log_area.pack(fill='both', expand=True)

    def choose_file(self):
        file = filedialog.askopenfilename(filetypes=[('Outlook Files', '*.pst *.ost')])
        if file:
            self.file_path.set(file)

    def log(self, msg):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, msg + '\n')
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def start_repair(self):
        pst = self.file_path.get()
        office = self.office_version.get()
        if not office or office not in OFFICE_PATHS:
            messagebox.showerror("Lỗi", "Vui lòng chọn đúng phiên bản Office.")
            return
        if not os.path.isfile(pst):
            messagebox.showerror("Lỗi", "Vui lòng chọn file PST/OST hợp lệ.")
            return
        exe = OFFICE_PATHS[office]
        threading.Thread(target=self.run_repair, args=(exe, pst), daemon=True).start()

    def run_repair(self, exe_path, pst_path):
        self.log("=== Bắt đầu sửa lỗi ===")
        if not PYWINAUTO_AVAILABLE:
            self.log("[LỖI] pywinauto chưa được cài. Không thể tự động hóa GUI.")
            return
        try:
            app = Application(backend='win32').start(exe_path)
            dlg = app.window(title_re=".*Inbox Repair Tool.*")
            dlg.wait('visible', timeout=20)
            self.log("[✓] Đã mở SCANPST.EXE")

            dlg.Edit.set_edit_text(pst_path)
            start_btn = dlg.child_window(title="Start", class_name="Button")
            start_btn.wait('enabled', timeout=10)
            start_btn.click()
            self.log("[✓] Đã bắt đầu quét file")

            time.sleep(10)
            repair_btn = dlg.child_window(title="Repair", class_name="Button")
            if repair_btn.exists(timeout=5):
                repair_btn.click()
                self.log("[✓] Đang sửa lỗi...")
                time.sleep(10)
                try:
                    ok_btn = dlg.child_window(title="OK", class_name="Button")
                    if ok_btn.exists():
                        ok_btn.click()
                        self.log("[✓] Sửa lỗi hoàn tất!")
                except:
                    self.log("[-] Không tìm thấy nút OK")
            else:
                try:
                    ok_btn = dlg.child_window(title="OK", class_name="Button")
                    if ok_btn.exists():
                        ok_btn.click()
                        self.log("[✓] Không phát hiện lỗi trong file")
                except:
                    self.log("[-] Không tìm thấy nút OK")
        except Exception as e:
            self.log(f"[LỖI] {e}")

if __name__ == '__main__':
    if not PYWINAUTO_AVAILABLE:
        print("Lỗi: Cần cài pywinauto. Chạy: pip install pywinauto")
    else:
        root = tk.Tk()
        app = ScanPstLiteWin7(root)
        root.mainloop()
