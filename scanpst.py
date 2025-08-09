import os
import time
import threading
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, scrolledtext

try:
    from pywinauto.application import Application
    from pywinauto.findwindows import ElementNotFoundError
    from pywinauto.timings import TimeoutError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

OFFICE_PATHS = {
    "Office 365/2019/2016 (64-bit)": r"C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE",
    "Office 365/2019/2016 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE",
    "Office 2013 (64-bit)": r"C:\Program Files\Microsoft Office\Office15\SCANPST.EXE",
    "Office 2013 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\Office15\SCANPST.EXE",
    "Office 2010 (64-bit)": r"C:\Program Files\Microsoft Office\Office14\SCANPST.EXE",
    "Office 2010 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\Office14\SCANPST.EXE",
}

class ScanPstApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Công Cụ Tự Động Sửa File PST/OST @danhcp")
        self.root.geometry("750x600")

        self.files_to_scan = []
        self.file_selection_info = tk.StringVar(value="Chưa chọn file nào")
        self.create_backup = tk.BooleanVar(value=False)
        self.selected_office_var = tk.StringVar()

        main_container = ttk.Frame(root, padding=10)
        main_container.pack(fill=BOTH, expand=YES)

        # <<< THAY ĐỔI: TÁI CẤU TRÚC LAYOUT >>>

        # 1. KHUNG CẤU HÌNH (Phía trên)
        settings_frame = ttk.LabelFrame(main_container, text=" Cấu Hình ", padding=15, bootstyle="primary")
        settings_frame.pack(side=TOP, fill=X, pady=(0, 10))

        label1 = ttk.Label(settings_frame, text="1. Chọn phiên bản Office của bạn:")
        label1.grid(row=0, column=0, columnspan=2, sticky=W, pady=(0, 5))
        self.office_combobox = ttk.Combobox(settings_frame, textvariable=self.selected_office_var, state='readonly', bootstyle="primary")
        self.office_combobox.grid(row=1, column=0, columnspan=2, sticky=EW, pady=(0, 15))
        
        label2 = ttk.Label(settings_frame, text="2. Chọn các file PST/OST cần sửa:")
        label2.grid(row=2, column=0, columnspan=2, sticky=W, pady=(0, 5))
        self.select_files_button = ttk.Button(settings_frame, text="Chọn Files...", command=self.select_files, bootstyle="secondary-outline")
        self.select_files_button.grid(row=3, column=0, sticky=W)
        self.selection_label = ttk.Label(settings_frame, textvariable=self.file_selection_info, bootstyle="secondary")
        self.selection_label.grid(row=3, column=1, sticky=W, padx=(10, 0), pady=(0, 15))
        settings_frame.columnconfigure(1, weight=1)

        self.backup_check = ttk.Checkbutton(settings_frame, text="Tạo file backup (.bak) trước khi sửa", variable=self.create_backup, bootstyle="primary-round-toggle")
        self.backup_check.grid(row=4, column=0, columnspan=2, sticky=W, pady=(10, 0))

        # <<< THAY ĐỔI: THÊM SPINBOX CHO SỐ LẦN LẶP >>>
        self.loop_count = tk.IntVar(value=1)
        loop_label = ttk.Label(settings_frame, text="3. Số lần lặp lại quá trình sửa (1-3):")
        loop_label.grid(row=5, column=0, columnspan=2, sticky=W, pady=(15, 5))
        self.loop_spinbox = ttk.Spinbox(settings_frame, from_=1, to=10, textvariable=self.loop_count, state='readonly', width=8, bootstyle="primary")
        self.loop_spinbox.grid(row=6, column=0, sticky=W)
        # <<< KẾT THÚC THAY ĐỔI >>>

        # 2. KHUNG HÀNH ĐỘNG (Ở giữa, ngay dưới cấu hình)
        action_frame = ttk.Frame(main_container)
        action_frame.pack(side=TOP, fill=X, pady=(5, 10))
        self.repair_button = ttk.Button(
            action_frame, text="BẮT ĐẦU SỬA LỖI", command=self.start_repair_thread, bootstyle="primary"
        )
        self.repair_button.pack(fill=X, ipady=5)

        # 3. KHUNG LOG (Dưới cùng, chiếm toàn bộ không gian còn lại)
        log_frame = ttk.LabelFrame(main_container, text=" Trạng Thái Hoạt Động ", padding=15, bootstyle="info")
        log_frame.pack(side=TOP, fill=BOTH, expand=YES)
        
        self.log_widget = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', font=("Segoe UI", 9), relief=FLAT)
        self.log_widget.pack(fill=BOTH, expand=YES)
        
        self.log_widget.tag_config('success', foreground='#198754', font=('Segoe UI', 9, 'bold'))
        self.log_widget.tag_config('error', foreground='#dc3545', font=('Segoe UI', 9, 'bold'))
        self.log_widget.tag_config('info', foreground='#0d6efd')
        self.log_widget.tag_config('header', foreground='#212529', font=('Segoe UI', 9, 'bold'))
        self.log_widget.tag_config('step', foreground='#6c757d')

        self.populate_office_combobox()

    # --- Các hàm logic không thay đổi ---
    def log(self, message, level="normal"):
        def _log(): self.log_widget.config(state='normal'); self.log_widget.insert(tk.END, message + "\n", level); self.log_widget.config(state='disabled'); self.log_widget.see(tk.END)
        self.root.after(0, _log)
        
    def select_files(self):
        file_types = [('Outlook Data Files', '*.pst *.ost'), ('All files', '*.*')]
        selected_files = filedialog.askopenfilenames(title="Chọn các file PST/OST cần sửa", filetypes=file_types)
        if selected_files:
            self.files_to_scan = list(selected_files)
            count = len(self.files_to_scan)
            self.file_selection_info.set(f"Đã chọn {count} file")
            self.log(f"[INFO] Đã chọn {count} file để xử lý:", level='info')
            for f in self.files_to_scan: self.log(f"  -> {os.path.basename(f)}")
        else: self.log("[INFO] Hộp thoại chọn file đã đóng.", level='info')

    def start_repair_thread(self):
        selected_key = self.selected_office_var.get()
        if not selected_key: messagebox.showerror("Lỗi", "Không có phiên bản Office hợp lệ nào để chọn."); return
        if not self.files_to_scan: messagebox.showerror("Lỗi", "Bạn chưa chọn file PST/OST nào để sửa."); return
        scanpst_path = OFFICE_PATHS[selected_key]
        self.set_controls_state("disabled")
        # <<< THAY ĐỔI: TRUYỀN SỐ LẦN LẶP VÀO THREAD >>>
        loop_count = self.loop_count.get()
        repair_thread = threading.Thread(target=self.run_repair_logic, args=(scanpst_path, list(self.files_to_scan), self.create_backup.get(), loop_count))
        repair_thread.daemon = True; repair_thread.start()

    def run_repair_logic(self, scanpst_path, files_to_process, create_backup, loop_count):
        try:
            self.log("="*50, level='header')
            self.log(f"Bắt đầu quá trình với {len(files_to_process)} file. Số lần lặp: {loop_count}", level='info')
            if not PYWINAUTO_AVAILABLE: self.log("[LỖI] Thư viện 'pywinauto' chưa được cài đặt.", level='error'); return
            
            success_count, fail_count = 0, 0
            # <<< THAY ĐỔI: VÒNG LẶP LỚN CHO TOÀN BỘ QUÁ TRÌNH >>>
            for loop_num in range(loop_count):
                self.log(f"\n==================== LƯỢT SỬA {loop_num + 1}/{loop_count} ====================", level='header')
                # Reset đếm cho mỗi lượt nếu muốn theo dõi riêng, hoặc giữ nguyên để đếm tổng
                
                for i, file_path in enumerate(files_to_process):
                    self.log(f"\n--- [Lượt {loop_num+1}] Xử lý file {i+1}/{len(files_to_process)} ---", level='header')
                    if self.repair_single_file(file_path, scanpst_path, create_backup): 
                        if loop_num == loop_count - 1: success_count += 1 # Chỉ đếm ở lượt cuối
                    else: 
                        if loop_num == loop_count - 1: fail_count += 1 # Chỉ đếm ở lượt cuối
            
            self.log("\n================ TỔNG KẾT CUỐI CÙNG ================", level='header')
            self.log(f"Hoàn thành: {success_count} file.", level='success')
            self.log(f"Thất bại/Bỏ qua: {fail_count} file.", level='error' if fail_count > 0 else 'step')

        except Exception as e: self.log(f"[LỖI NGHIÊM TRỌNG] Lỗi không mong muốn: {e}", level='error')
        finally: self.set_controls_state("normal")

    def populate_office_combobox(self):
        valid_options = [name for name, path in OFFICE_PATHS.items() if os.path.exists(path)]
        self.office_combobox['values'] = valid_options
        if valid_options:
            self.selected_office_var.set(valid_options[0])
            self.log("Chào mừng bạn!", level='header'); self.log("[INFO] Đã tìm thấy các phiên bản Office hợp lệ.", level='info')
        else:
            self.log("[LỖI] Không tìm thấy file SCANPST.EXE.", level='error'); self.repair_button.config(state="disabled")

    def set_controls_state(self, state):
        combobox_state = 'readonly' if state == "normal" else 'disabled'
        self.repair_button.config(state=state); self.select_files_button.config(state=state)
        self.office_combobox.config(state=combobox_state); self.backup_check.config(state=state)
        # <<< THAY ĐỔI: VÔ HIỆU HÓA SPINBOX KHI CHẠY >>>
        self.loop_spinbox.config(state=state)
        
    def repair_single_file(self, pst_path, scanpst_path, create_backup):
        self.log(f"[*] Đang xử lý file: {os.path.basename(pst_path)}", level='info')
        app = None
        try:
            app = Application(backend="uia").start(scanpst_path)
            dlg = app.window(title="Microsoft Outlook Inbox Repair Tool")
            dlg.wait('visible', timeout=30)
            self.log("    [1] Đã mở ScanPST.", level='step')
            try:
                dlg.ComboBox.set_edit_text(pst_path)
            except Exception:
                dlg.Edit.set_edit_text(pst_path)
            self.log("    [2] Đã nhập đường dẫn file.", level='step')
            
            dlg.child_window(title="Start", control_type="Button").click()
            self.log("    [3] Đã nhấn 'Start'. Đang quét file...", level='step')

            scan_timeout = 12000; start_time = time.time(); scan_finished = False
            while time.time() - start_time < scan_timeout:
                repair_button = dlg.child_window(title="Repair", control_type="Button")
                if repair_button.exists() and repair_button.is_enabled():
                    self.log("    [4] Quét xong. Phát hiện có lỗi.", level='info')
                    backup_checkbox = dlg.child_window(title_re=".*Make backup.*", control_type="CheckBox")
                    if backup_checkbox.exists():
                        if not create_backup:
                            backup_checkbox.click_input(); self.log("    [5] Đã bỏ tick ô backup.", level='step')
                        else: self.log("    [5] Giữ nguyên tùy chọn backup.", level='step')
                    repair_button.click()
                    self.log("    [6] Đã nhấn 'Repair'. Bắt đầu sửa lỗi...", level='step')
                    repair_timeout = scan_timeout - (time.time() - start_time); repair_start_time = time.time()
                    while time.time() - repair_start_time < repair_timeout:
                        try:
                            popup = app.window(title="Microsoft Outlook Inbox Repair Tool", top_level_only=True)
                            if popup.child_window(title="Yes").exists():
                                popup.Yes.click_input(); self.log("    [+] Backup đã tồn tại, đã nhấn 'Yes'.", level='info')
                                time.sleep(1); continue
                            if popup.child_window(title="OK").exists():
                                popup.OK.click_input(); self.log("    [7] Sửa lỗi hoàn tất!", level='success')
                                scan_finished = True; break
                        except (TimeoutError, ElementNotFoundError): pass
                        time.sleep(0.5)
                    if scan_finished: break
                    else: raise TimeoutError("Quá trình sửa lỗi quá thời gian.")

                ok_button = dlg.child_window(title="OK", control_type="Button")
                if ok_button.exists() and ok_button.is_enabled():
                    self.log("    [4] Quét xong. Không tìm thấy lỗi.", level='success')
                    ok_button.click_input(); scan_finished = True; break
                time.sleep(0.5)
            
            if not scan_finished: raise TimeoutError("Quá trình quét/sửa lỗi quá thời gian.")
            if dlg.exists(): dlg.close()
            return True

        except (TimeoutError, ElementNotFoundError) as e:
            self.log(f"    [LỖI] Timeout hoặc không tìm thấy cửa sổ/nút bấm: {e}", level='error')
            if app and app.is_process_running(): app.kill(); return False
        except Exception as e:
            self.log(f"    [LỖI] Lỗi không xác định: {e}", level='error')
            if app and app.is_process_running(): app.kill(); return False

if __name__ == "__main__":
    if not PYWINAUTO_AVAILABLE:
        print("Lỗi: Thư viện 'pywinauto' chưa được cài đặt."); input("Nhấn Enter để thoát.")
    else:
        main_root = ttk.Window(themename="lumen")
        app = ScanPstApp(main_root)
        main_root.mainloop()