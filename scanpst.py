import os
import sys
import time
import winreg
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

try:
    from pywinauto.application import Application
    from pywinauto.findwindows import ElementNotFoundError
    from pywinauto.timings import TimeoutError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


# Danh sách đường dẫn mặc định theo các phiên bản Office phổ biến
DEFAULT_OFFICE_PATHS = {
    "Office 365/2019/2016 (64-bit)": r"C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE",
    "Office 365/2019/2016 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE",
    "Office 2013 (64-bit)": r"C:\Program Files\Microsoft Office\Office15\SCANPST.EXE",
    "Office 2013 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\Office15\SCANPST.EXE",
    "Office 2010 (64-bit)": r"C:\Program Files\Microsoft Office\Office14\SCANPST.EXE",
    "Office 2010 (32-bit)": r"C:\Program Files (x86)\Microsoft Office\Office14\SCANPST.EXE",
}


def is_process_running(process_name: str) -> bool:
    """Kiểm tra xem một tiến trình có đang chạy trong Windows hay không."""
    try:
        output = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {process_name}", "/NH"],
            text=True,
            errors="ignore"
        )
        return process_name.lower() in output.lower()
    except Exception:
        return False


def kill_process(process_name: str) -> bool:
    """Tắt cưỡng bức tiến trình theo tên."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", process_name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False
        )
        return True
    except Exception:
        return False


def find_scanpst_from_registry() -> dict:
    """
    Truy vấn đường dẫn SCANPST.EXE từ Windows Registry:
    1. App Paths của OUTLOOK.EXE (cả 32-bit và 64-bit).
    2. InstallRoot của Microsoft Office trong HKLM.
    """
    discovered_paths = {}

    # 1. Truy vấn qua App Paths
    app_path_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
        (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
    ]

    for root, subkey in app_path_keys:
        try:
            with winreg.OpenKey(root, subkey) as key:
                outlook_exe_path, _ = winreg.QueryValueEx(key, "")
                if outlook_exe_path and os.path.exists(outlook_exe_path):
                    office_dir = os.path.dirname(outlook_exe_path)
                    scanpst_candidate = os.path.join(office_dir, "SCANPST.EXE")
                    if os.path.exists(scanpst_candidate):
                        norm_path = os.path.normpath(scanpst_candidate)
                        name = f"Outlook Registry ({os.path.basename(office_dir)})"
                        discovered_paths[name] = norm_path
        except OSError:
            pass

    # 2. Truy vấn qua InstallRoot của Office
    office_versions = [
        ("16.0", "Office 365/2016/2019/2021/2024"),
        ("15.0", "Office 2013"),
        ("14.0", "Office 2010"),
        ("12.0", "Office 2007"),
    ]

    for ver_code, ver_name in office_versions:
        reg_subkeys = [
            f"SOFTWARE\\Microsoft\\Office\\{ver_code}\\Outlook\\InstallRoot",
            f"SOFTWARE\\WOW6432Node\\Microsoft\\Office\\{ver_code}\\Outlook\\InstallRoot",
        ]
        for subkey in reg_subkeys:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey) as key:
                    install_path, _ = winreg.QueryValueEx(key, "Path")
                    if install_path and os.path.exists(install_path):
                        scanpst_candidate = os.path.join(install_path, "SCANPST.EXE")
                        if os.path.exists(scanpst_candidate):
                            norm_path = os.path.normpath(scanpst_candidate)
                            name = f"{ver_name} (Registry)"
                            if norm_path not in discovered_paths.values():
                                discovered_paths[name] = norm_path
            except OSError:
                pass

    return discovered_paths


class ScanPstApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Công Cụ Tự Động Sửa File PST/OST @danhcp")
        self.root.geometry("780x640")

        # Cấu hình icon cho cửa sổ
        try:
            icon_path = self.get_icon_path()
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        self.files_to_scan = []
        self.file_selection_info = tk.StringVar(value="Chưa chọn file nào")
        self.create_backup = tk.BooleanVar(value=False)
        self.selected_office_var = tk.StringVar()
        self.loop_count = tk.IntVar(value=1)

        # Quản lý đường dẫn SCANPST và trạng thái dừng
        self.available_office_paths = {}
        self.stop_requested = False
        self.is_running = False
        self.current_app = None

        self._build_ui()
        self.populate_office_combobox()

    def _build_ui(self):
        """Khởi tạo toàn bộ giao diện người dùng."""
        main_container = ttk.Frame(self.root, padding=12)
        main_container.pack(fill=BOTH, expand=YES)

        # 1. KHUNG CẤU HÌNH (Phía trên)
        settings_frame = ttk.Labelframe(
            main_container,
            text=" Cấu Hình Hoạt Động ",
            padding=15,
            bootstyle="primary"
        )
        settings_frame.pack(side=TOP, fill=X, pady=(0, 10))

        # 1.1. Chọn SCANPST.EXE
        lbl_office = ttk.Label(settings_frame, text="1. Phiên bản Microsoft Office / SCANPST.EXE:")
        lbl_office.grid(row=0, column=0, columnspan=3, sticky=W, pady=(0, 5))

        self.office_combobox = ttk.Combobox(
            settings_frame,
            textvariable=self.selected_office_var,
            state="readonly",
            bootstyle="primary"
        )
        self.office_combobox.grid(row=1, column=0, columnspan=2, sticky=EW, pady=(0, 15))

        self.btn_browse_scanpst = ttk.Button(
            settings_frame,
            text="Duyệt File...",
            command=self.browse_scanpst_manually,
            bootstyle="outline-primary"
        )
        self.btn_browse_scanpst.grid(row=1, column=2, sticky=W, padx=(8, 0), pady=(0, 15))

        # 1.2. Chọn file PST/OST
        lbl_files = ttk.Label(settings_frame, text="2. Danh sách file PST/OST cần sửa:")
        lbl_files.grid(row=2, column=0, columnspan=3, sticky=W, pady=(0, 5))

        self.select_files_button = ttk.Button(
            settings_frame,
            text="Chọn Files...",
            command=self.select_files,
            bootstyle="secondary-outline"
        )
        self.select_files_button.grid(row=3, column=0, sticky=W)

        self.selection_label = ttk.Label(
            settings_frame,
            textvariable=self.file_selection_info,
            bootstyle="secondary"
        )
        self.selection_label.grid(row=3, column=1, columnspan=2, sticky=W, padx=(10, 0))

        # 1.3. Tuỳ chọn Backup
        self.backup_check = ttk.Checkbutton(
            settings_frame,
            text="Tạo file backup (.bak) trước khi sửa",
            variable=self.create_backup,
            bootstyle="primary-round-toggle"
        )
        self.backup_check.grid(row=4, column=0, columnspan=3, sticky=W, pady=(15, 5))

        # 1.4. Số lần lặp lại
        lbl_loop = ttk.Label(settings_frame, text="3. Số lần lặp lại quá trình sửa (1 - 10 lần):")
        lbl_loop.grid(row=5, column=0, columnspan=3, sticky=W, pady=(10, 5))

        self.loop_spinbox = ttk.Spinbox(
            settings_frame,
            from_=1,
            to=10,
            textvariable=self.loop_count,
            state="readonly",
            width=8,
            bootstyle="primary"
        )
        self.loop_spinbox.grid(row=6, column=0, sticky=W)

        settings_frame.columnconfigure(1, weight=1)

        # 2. KHUNG HÀNH ĐỘNG (Nút Bắt đầu & Nút Dừng)
        action_frame = ttk.Frame(main_container)
        action_frame.pack(side=TOP, fill=X, pady=(5, 10))

        self.repair_button = ttk.Button(
            action_frame,
            text="▶  BẮT ĐẦU SỬA LỖI",
            command=self.start_repair_thread,
            bootstyle="primary"
        )
        self.repair_button.pack(side=LEFT, fill=X, expand=YES, padx=(0, 6), ipady=6)

        self.stop_button = ttk.Button(
            action_frame,
            text="⏹  DỪNG LẠI",
            command=self.stop_repair,
            bootstyle="danger",
            state="disabled"
        )
        self.stop_button.pack(side=RIGHT, fill=X, expand=NO, ipady=6)

        # 3. KHUNG LOG HOẠT ĐỘNG
        log_frame = ttk.Labelframe(
            main_container,
            text=" Trạng Thái Hoạt Động ",
            padding=12,
            bootstyle="info"
        )
        log_frame.pack(side=TOP, fill=BOTH, expand=YES)

        self.log_widget = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            state="disabled",
            font=("Segoe UI", 9),
            relief=FLAT
        )
        self.log_widget.pack(fill=BOTH, expand=YES)

        self.log_widget.tag_config("success", foreground="#198754", font=("Segoe UI", 9, "bold"))
        self.log_widget.tag_config("error", foreground="#dc3545", font=("Segoe UI", 9, "bold"))
        self.log_widget.tag_config("info", foreground="#0d6efd")
        self.log_widget.tag_config("header", foreground="#212529", font=("Segoe UI", 9, "bold"))
        self.log_widget.tag_config("step", foreground="#6c757d")
        self.log_widget.tag_config("warning", foreground="#fd7e14", font=("Segoe UI", 9, "bold"))

    def get_icon_path(self) -> str:
        """Lấy đường dẫn icon từ thư mục hiện tại hoặc từ bundle PyInstaller."""
        if getattr(sys, "frozen", False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, "icon.ico")

    def log(self, message: str, level: str = "normal"):
        """Ghi nội dung log vào ScrolledText an toàn luồng."""
        def _append_log():
            self.log_widget.config(state="normal")
            self.log_widget.insert(tk.END, message + "\n", level)
            self.log_widget.config(state="disabled")
            self.log_widget.see(tk.END)

        self.root.after(0, _append_log)

    def populate_office_combobox(self):
        """Dò tìm SCANPST.EXE từ Windows Registry và đường dẫn mặc định."""
        self.available_office_paths = {}

        # 1. Truy vấn từ Registry
        registry_paths = find_scanpst_from_registry()
        for name, path in registry_paths.items():
            self.available_office_paths[f"[Registry] {name}"] = path

        # 2. Truy vấn từ các đường dẫn mặc định
        for name, path in DEFAULT_OFFICE_PATHS.items():
            if os.path.exists(path):
                norm_path = os.path.normpath(path)
                if norm_path not in self.available_office_paths.values():
                    self.available_office_paths[name] = norm_path

        valid_keys = list(self.available_office_paths.keys())
        self.office_combobox["values"] = valid_keys

        if valid_keys:
            self.selected_office_var.set(valid_keys[0])
            self.log("Chào mừng bạn đến với Công Cụ Tự Động Sửa File PST/OST!", level="header")
            self.log(f"[INFO] Đã tìm thấy {len(valid_keys)} vị trí SCANPST.EXE khả dụng.", level="info")
        else:
            self.log("[CẢNH BÁO] Không tìm thấy SCANPST.EXE tự động. Vui lòng bấm 'Duyệt File...' để chọn thủ công.", level="warning")
            self.repair_button.config(state="disabled")

    def browse_scanpst_manually(self):
        """Cho phép người dùng tự chọn file SCANPST.EXE bằng tay."""
        file_path = filedialog.askopenfilename(
            title="Chọn file SCANPST.EXE",
            filetypes=[("SCANPST Executable", "SCANPST.EXE"), ("All Executables", "*.exe"), ("All Files", "*.*")]
        )
        if file_path and os.path.exists(file_path):
            norm_path = os.path.normpath(file_path)
            key_name = f"[Tùy chọn] {norm_path}"
            self.available_office_paths[key_name] = norm_path
            current_values = list(self.office_combobox["values"])
            if key_name not in current_values:
                current_values.append(key_name)
                self.office_combobox["values"] = current_values
            self.selected_office_var.set(key_name)
            self.repair_button.config(state="normal")
            self.log(f"[INFO] Đã chọn đường dẫn SCANPST thủ công: {norm_path}", level="info")

    def select_files(self):
        """Mở hộp thoại chọn file PST/OST cần sửa."""
        file_types = [
            ("Outlook Data Files", "*.pst *.ost"),
            ("PST Files (*.pst)", "*.pst"),
            ("OST Files (*.ost)", "*.ost"),
            ("All files", "*.*")
        ]
        selected_files = filedialog.askopenfilenames(
            title="Chọn các file PST/OST cần sửa",
            filetypes=file_types
        )
        if selected_files:
            self.files_to_scan = list(selected_files)
            count = len(self.files_to_scan)
            self.file_selection_info.set(f"Đã chọn {count} file")
            self.log(f"[INFO] Đã chọn {count} file để xử lý:", level="info")
            for file_path in self.files_to_scan:
                self.log(f"  -> {os.path.basename(file_path)}")
        else:
            self.log("[INFO] Hộp thoại chọn file đã đóng.", level="info")

    def set_controls_state(self, running: bool):
        """Cập nhật trạng thái các nút điều khiển khi bắt đầu hoặc kết thúc."""
        self.is_running = running
        if running:
            self.repair_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.select_files_button.config(state="disabled")
            self.btn_browse_scanpst.config(state="disabled")
            self.office_combobox.config(state="disabled")
            self.backup_check.config(state="disabled")
            self.loop_spinbox.config(state="disabled")
        else:
            self.repair_button.config(state="normal")
            self.stop_button.config(state="disabled")
            self.select_files_button.config(state="normal")
            self.btn_browse_scanpst.config(state="normal")
            self.office_combobox.config(state="readonly")
            self.backup_check.config(state="normal")
            self.loop_spinbox.config(state="readonly")

    def check_and_close_outlook(self) -> bool:
        """
        Kiểm tra nếu Outlook hoặc SCANPST cũ đang mở:
        - Nhắc người dùng tự động đóng để tránh khóa file độc quyền (file lock).
        """
        if is_process_running("OUTLOOK.EXE"):
            ans = messagebox.askyesno(
                "Phát hiện Outlook đang chạy",
                "Microsoft Outlook đang mở! Để quét và sửa chữa file PST/OST, "
                "Outlook cần phải được đóng hoàn toàn để tránh xung đột khóa file.\n\n"
                "Bạn có muốn chương trình tự động đóng Microsoft Outlook ngay bây giờ không?"
            )
            if ans:
                self.log("[*] Đang đóng Microsoft Outlook...", level="info")
                kill_process("OUTLOOK.EXE")
                time.sleep(2)
                if is_process_running("OUTLOOK.EXE"):
                    messagebox.showerror("Lỗi", "Không thể đóng Microsoft Outlook. Vui lòng đóng Outlook thủ công rồi thử lại.")
                    return False
                self.log("[OK] Đã đóng Microsoft Outlook thành công.", level="success")
            else:
                self.log("[HỦY] Người dùng từ chối đóng Outlook. Quá trình dừng lại.", level="warning")
                return False

        # Dọn dẹp tiến trình SCANPST.EXE còn sót lại từ các lần chạy trước
        if is_process_running("SCANPST.EXE"):
            self.log("[*] Đang dọn dẹp các tiến trình SCANPST.EXE cũ còn chạy ngầm...", level="step")
            kill_process("SCANPST.EXE")
            time.sleep(1)

        return True

    def start_repair_thread(self):
        """Khởi động luồng xử lý quét và sửa lỗi."""
        selected_key = self.selected_office_var.get()
        if not selected_key or selected_key not in self.available_office_paths:
            messagebox.showerror("Lỗi", "Vui lòng chọn hoặc duyệt tìm file SCANPST.EXE hợp lệ.")
            return

        if not self.files_to_scan:
            messagebox.showerror("Lỗi", "Bạn chưa chọn file PST/OST nào để sửa.")
            return

        if not PYWINAUTO_AVAILABLE:
            messagebox.showerror("Lỗi", "Thư viện 'pywinauto' chưa được cài đặt. Không thể tự động hóa giao diện.")
            return

        # Kiểm tra và đóng Outlook trước khi tiến hành
        if not self.check_and_close_outlook():
            return

        scanpst_path = self.available_office_paths[selected_key]
        loop_count = self.loop_count.get()
        create_backup = self.create_backup.get()

        self.stop_requested = False
        self.set_controls_state(running=True)

        repair_thread = threading.Thread(
            target=self.run_repair_logic,
            args=(scanpst_path, list(self.files_to_scan), create_backup, loop_count),
            daemon=True
        )
        repair_thread.start()

    def stop_repair(self):
        """Yêu cầu dừng tiến trình đang chạy."""
        if not self.is_running:
            return

        self.stop_requested = True
        self.stop_button.config(state="disabled")
        self.log("\n[CẢNH BÁO] Đang gửi yêu cầu dừng tiến trình, vui lòng chờ trong giây lát...", level="warning")

        # Đóng tiến trình SCANPST hiện tại nếu đang mở
        if self.current_app and self.current_app.is_process_running():
            try:
                self.current_app.kill()
            except Exception:
                pass

    def run_repair_logic(self, scanpst_path, files_to_process, create_backup, loop_count):
        """Toàn bộ logic lặp và quét danh sách file qua luồng phụ."""
        try:
            self.log("=" * 60, level="header")
            self.log(
                f"BẮT ĐẦU QUÁ TRÌNH: {len(files_to_process)} file | Số lượt lặp: {loop_count} | Backup: {'Bật' if create_backup else 'Tắt'}",
                level="info"
            )
            self.log(f"Đường dẫn SCANPST: {scanpst_path}", level="step")
            self.log("=" * 60, level="header")

            success_count = 0
            fail_count = 0

            for loop_num in range(loop_count):
                if self.stop_requested:
                    self.log("[DỪNG] Người dùng đã dừng quá trình lặp.", level="warning")
                    break

                self.log(f"\n==================== LƯỢT SỬA {loop_num + 1}/{loop_count} ====================", level="header")

                for i, file_path in enumerate(files_to_process):
                    if self.stop_requested:
                        self.log("[DỪNG] Quá trình xử lý file đã bị người dùng hủy.", level="warning")
                        break

                    self.log(
                        f"\n--- [Lượt {loop_num + 1}/{loop_count}] File {i + 1}/{len(files_to_process)}: {os.path.basename(file_path)} ---",
                        level="header"
                    )

                    is_ok = self.repair_single_file(file_path, scanpst_path, create_backup)

                    # Chỉ ghi nhận thống kê ở lượt lặp cuối cùng
                    if loop_num == loop_count - 1:
                        if is_ok:
                            success_count += 1
                        else:
                            fail_count += 1

            self.log("\n================ TỔNG KẾT TIẾN TRÌNH ================", level="header")
            if self.stop_requested:
                self.log("Trạng thái: ĐÃ DỪNG BỞI NGƯỜI DÙNG", level="warning")
            else:
                self.log(f"Thành công: {success_count} file.", level="success")
                if fail_count > 0:
                    self.log(f"Thất bại / Bỏ qua: {fail_count} file.", level="error")
                else:
                    self.log("Tất cả các file đều được xử lý hoàn tất!", level="success")

        except Exception as e:
            self.log(f"[LỖI NGHIÊM TRỌNG] Ngoại lệ ngoài ý muốn: {e}", level="error")
        finally:
            self.current_app = None
            self.set_controls_state(running=False)

    def _sync_checkbox_state(self, checkbox, desired_state: bool):
        """
        Đồng bộ trạng thái của Checkbox sao cho đúng với mong muốn (desired_state).
        Tránh click mù quáng làm đảo ngược trạng thái.
        """
        current_state = None
        try:
            # Đối với UIA backend, get_toggle_state() trả về 1 nếu Checked, 0 nếu Unchecked
            current_state = bool(checkbox.get_toggle_state())
        except Exception:
            try:
                # Fallback qua MSAA/legacy properties
                props = checkbox.legacy_properties()
                state_val = props.get("State", "")
                if isinstance(state_val, str):
                    current_state = "checked" in state_val.lower()
            except Exception:
                pass

        if current_state is not None:
            if current_state != desired_state:
                checkbox.click_input()
                self.log(
                    f"    [+] Đã chuyển trạng thái Backup checkbox thành: {'BẬT' if desired_state else 'TẮT'}.",
                    level="step"
                )
            else:
                self.log(
                    f"    [+] Backup checkbox đã ở trạng thái mong muốn: {'BẬT' if desired_state else 'TẮT'}.",
                    level="step"
                )
        else:
            # Nếu không thể đọc được trạng thái, chỉ click khi muốn TẮT (vì mặc định SCANPST là BẬT)
            if not desired_state:
                checkbox.click_input()
                self.log("    [+] Đã click bỏ chọn ô backup.", level="step")

    def repair_single_file(self, pst_path: str, scanpst_path: str, create_backup: bool) -> bool:
        """
        Tự động hóa toàn bộ quy trình SCANPST.EXE cho một file duy nhất.
        LƯU Ý QUAN TRỌNG:
        - KHÔNG giới hạn thời gian (No Hard Timeout): Các file lớn hoặc lỗi nặng có thể quét rất lâu.
        - Giám sát tiến trình liên tục: nếu SCANPST.EXE bị crash/tắt đột ngột thì mới báo lỗi và thoát.
        """
        self.log(f"[*] Bắt đầu xử lý: {pst_path}", level="info")
        self.current_app = None

        if not os.path.exists(pst_path):
            self.log(f"    [LỖI] File không tồn tại trên đĩa: {pst_path}", level="error")
            return False

        try:
            self.current_app = Application(backend="uia").start(scanpst_path)
            dlg = self.current_app.window(title="Microsoft Outlook Inbox Repair Tool")
            dlg.wait("visible", timeout=30)
            self.log("    [1] Đã khởi chạy SCANPST.EXE.", level="step")

            if self.stop_requested:
                dlg.close()
                return False

            # Nhập đường dẫn file PST/OST vào ô nhập liệu
            try:
                dlg.ComboBox.set_edit_text(pst_path)
            except Exception:
                dlg.Edit.set_edit_text(pst_path)
            self.log("    [2] Đã nhập đường dẫn file cần kiểm tra.", level="step")

            if self.stop_requested:
                dlg.close()
                return False

            # Nhấn nút Start
            dlg.child_window(title="Start", control_type="Button").click()
            self.log("    [3] Đã nhấn 'Start'. Đang quét phân tích cấu trúc file...", level="step")

            scan_finished = False
            start_time = time.time()
            last_heartbeat = start_time

            # Vòng lặp chờ kết quả quét (KHÔNG giới hạn thời gian, liên tục theo dõi tiến trình)
            while True:
                if self.stop_requested:
                    self.log("    [DỪNG] Huỷ xử lý theo yêu cầu người dùng.", level="warning")
                    if dlg.exists():
                        dlg.close()
                    return False

                # Kiểm tra nếu tiến trình SCANPST bị crash hoặc tắt ngoài ý muốn
                if not self.current_app.is_process_running():
                    self.log("    [LỖI] Tiến trình SCANPST.EXE đã bị tắt hoặc crash đột ngột.", level="error")
                    return False

                # Cập nhật thông báo tiến trình định kỳ mỗi 60 giây để người dùng yên tâm
                current_time = time.time()
                if current_time - last_heartbeat >= 60:
                    elapsed_mins = int((current_time - start_time) // 60)
                    self.log(
                        f"    [...] SCANPST đang tiếp tục quét/xử lý (đã chạy {elapsed_mins} phút), vui lòng chờ...",
                        level="step"
                    )
                    last_heartbeat = current_time

                # Trường hợp 1: Quét xong và phát hiện có lỗi -> Xuất hiện nút "Repair"
                repair_button = dlg.child_window(title="Repair", control_type="Button")
                if repair_button.exists() and repair_button.is_enabled():
                    self.log("    [4] Phân tích hoàn tất: Phát hiện file có lỗi cần sửa chữa.", level="info")

                    # Xử lý checkbox backup chính xác
                    backup_checkbox = dlg.child_window(title_re=".*Make backup.*", control_type="CheckBox")
                    if backup_checkbox.exists():
                        self._sync_checkbox_state(backup_checkbox, create_backup)

                    if self.stop_requested:
                        dlg.close()
                        return False

                    repair_button.click()
                    self.log("    [5] Đã nhấn 'Repair'. Đang sửa chữa các cấu trúc bị lỗi...", level="step")

                    repair_start_time = time.time()
                    repair_heartbeat = repair_start_time

                    # Vòng lặp chờ sửa chữa hoàn tất (KHÔNG giới hạn thời gian)
                    while True:
                        if self.stop_requested:
                            if dlg.exists():
                                dlg.close()
                            return False

                        if not self.current_app.is_process_running():
                            self.log("    [LỖI] Tiến trình SCANPST.EXE đã bị tắt trong khi đang sửa lỗi.", level="error")
                            return False

                        # Heartbeat cho giai đoạn sửa lỗi
                        now = time.time()
                        if now - repair_heartbeat >= 60:
                            elapsed_mins = int((now - repair_start_time) // 60)
                            self.log(
                                f"    [...] Quá trình sửa lỗi vẫn đang tiến hành (đã chạy {elapsed_mins} phút)...",
                                level="step"
                            )
                            repair_heartbeat = now

                        try:
                            popup = self.current_app.window(title="Microsoft Outlook Inbox Repair Tool", top_level_only=True)

                            # Popup hỏi ghi đè file backup nếu file .bak đã tồn tại
                            if popup.child_window(title="Yes").exists():
                                popup.Yes.click_input()
                                self.log("    [+] Đã xác nhận ghi đè file backup cũ ('Yes').", level="info")
                                time.sleep(1)
                                continue

                            # Popup thông báo hoàn tất "Repair complete"
                            if popup.child_window(title="OK").exists():
                                popup.OK.click_input()
                                self.log("    [6] Sửa lỗi hoàn tất thành công (Repair complete)!", level="success")
                                scan_finished = True
                                break

                        except (TimeoutError, ElementNotFoundError):
                            pass

                        time.sleep(0.5)

                    if scan_finished:
                        break

                # Trường hợp 2: Quét xong và không tìm thấy lỗi nào -> Nút "OK" xuất hiện
                ok_button = dlg.child_window(title="OK", control_type="Button")
                if ok_button.exists() and ok_button.is_enabled():
                    self.log("    [4] Phân tích hoàn tất: Không tìm thấy lỗi trong file này.", level="success")
                    ok_button.click_input()
                    scan_finished = True
                    break

                time.sleep(0.5)

            if dlg.exists():
                dlg.close()

            return True

        except (TimeoutError, ElementNotFoundError) as e:
            self.log(f"    [LỖI UI] Không tìm thấy cửa sổ hoặc phần tử điều khiển: {e}", level="error")
            if self.current_app and self.current_app.is_process_running():
                self.current_app.kill()
            return False

        except Exception as e:
            self.log(f"    [LỖI HỆ THỐNG] Lỗi ngoại lệ: {e}", level="error")
            if self.current_app and self.current_app.is_process_running():
                self.current_app.kill()
            return False

        finally:
            self.current_app = None


if __name__ == "__main__":
    if not PYWINAUTO_AVAILABLE:
        print("Lỗi: Thư viện 'pywinauto' chưa được cài đặt.")
        print("Vui lòng chạy lệnh: pip install pywinauto ttkbootstrap")
        input("Nhấn Enter để thoát...")
    else:
        main_root = ttk.Window(themename="lumen")
        app = ScanPstApp(main_root)
        main_root.mainloop()
