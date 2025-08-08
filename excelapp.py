# -*- coding: utf-8 -*-
"""
File: excelapp.py
Author: Your Name / Tên của bạn
Description: Chứa class ExcelApp để quản lý toàn bộ tiến trình Excel.

--- CHANGELOG ---
Version 0.3.0 (2025-08-08):
    - Thêm các tùy chọn tăng tốc vào __init__: screen_updating, display_alerts, calculation.
    - Thêm các thuộc tính tiện lợi: .workbooks, .workbook_names.
    - Thêm phương thức .kill_hidden_processes() để chỉ đóng các tiến trình Excel đang chạy ẩn.

Version 0.2.0 (2025-08-08):
    - Thêm phương thức .convert_to_xlsx() để chuyển đổi các định dạng file khác sang .xlsx.

Version 0.1.0 (2025-08-08):
    - Khởi tạo class ExcelApp với các tính năng cơ bản.
-------------------
"""

import xlwings as xw
import os
import time
from pathlib import Path
from .workbook import Workbook  # Sử dụng import tương đối

class ExcelApp:
    """
    Lớp quản lý chính, đại diện cho một tiến trình (instance) của ứng dụng Excel.
    """

    def __init__(self, visible=True, add_book=False, screen_updating=True, display_alerts=True, calculation='automatic'):
        """
        Khởi tạo và cấu hình ứng dụng Excel.

        Args:
            visible (bool): True để hiển thị cửa sổ Excel.
            add_book (bool): True để tự động tạo workbook mới.
            screen_updating (bool): True để Excel cập nhật màn hình. Tắt (False) để tăng tốc độ.
            display_alerts (bool): True để hiển thị cảnh báo của Excel. Tắt (False) để bỏ qua.
            calculation (str): Chế độ tính toán ('automatic', 'manual'). 'manual' giúp tăng tốc.
        """
        print("INFO: Khởi tạo tiến trình Excel...")
        try:
            self._app = xw.App(visible=visible, add_book=add_book)
            
            # Áp dụng các tùy chọn hiệu suất
            self._app.screen_updating = screen_updating
            self._app.display_alerts = display_alerts
            self._app.calculation = calculation
            
            print(f"INFO: Cấu hình Excel - ScreenUpdating: {screen_updating}, DisplayAlerts: {display_alerts}, Calculation: {calculation}")

        except Exception as e:
            print(f"ERROR: Không thể khởi tạo Excel. Vui lòng kiểm tra cài đặt của bạn. Lỗi: {e}")
            raise

    def __enter__(self):
        # Khi vào khối 'with', lưu lại trạng thái ban đầu
        self._initial_state = {
            'screen_updating': self._app.screen_updating,
            'display_alerts': self._app.display_alerts,
            'calculation': self._app.calculation
        }
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Khi thoát khối 'with', khôi phục lại trạng thái ban đầu trước khi đóng
        if self._app:
            try:
                self._app.screen_updating = self._initial_state['screen_updating']
                self._app.display_alerts = self._initial_state['display_alerts']
                self._app.calculation = self._initial_state['calculation']
            except Exception:
                pass # Bỏ qua nếu app đã bị lỗi
        self.quit()

    # --- Properties ---
    @property
    def workbooks(self):
        """Trả về một danh sách các đối tượng Workbook đang được quản lý."""
        return [Workbook(book, self) for book in self._app.books]

    @property
    def workbook_names(self):
        """Trả về một danh sách tên của các workbook đang mở."""
        return [book.name for book in self._app.books]

    # --- Methods ---
    def open(self, path, password=None, read_only=False):
        """Mở một workbook."""
        file_path = Path(path).resolve()
        for book in self._app.books:
            if Path(book.fullname).resolve() == file_path:
                return Workbook(book, self)
        try:
            xlw_book = self._app.books.open(file_path, password=password, read_only=read_only, ignore_read_only_recommended=True)
            return Workbook(xlw_book, self)
        except Exception as e:
            print(f"ERROR: Không thể mở workbook tại '{file_path}'. Lỗi: {e}")
            return None

    def new(self):
        """Tạo một workbook mới."""
        xlw_book = self._app.books.add()
        return Workbook(xlw_book, self)

    def get_workbook(self, specifier=None):
        """Lấy một workbook đã mở."""
        if not self._app.books: return None
        if specifier is None: return self.get_active_workbook()
        try:
            return Workbook(self._app.books[specifier], self)
        except Exception:
            return None

    def get_active_workbook(self):
        """Lấy workbook đang active."""
        if self._app.books.active:
            return Workbook(self._app.books.active, self)
        return None

    def wait_for_workbook(self, title_contains=None, title_is=None, timeout=30):
        """Chờ một workbook xuất hiện."""
        start_time = time.time()
        while time.time() - start_time < timeout:
            for book in self._app.books:
                if (title_is and book.name == title_is) or (title_contains and title_contains in book.name):
                    return Workbook(book, self)
            time.sleep(1)
        return None

    def convert_to_xlsx(self, source_path, destination_path=None):
        """Chuyển đổi file sang định dạng .xlsx."""
        source_path = Path(source_path)
        if not destination_path:
            destination_path = source_path.with_suffix('.xlsx')
        else:
            destination_path = Path(destination_path)
        
        temp_xlw_book = None
        try:
            temp_xlw_book = self._app.books.open(source_path)
            temp_xlw_book.save(destination_path)
            temp_xlw_book.close()
            return self.open(destination_path)
        except Exception as e:
            print(f"ERROR: Quá trình chuyển đổi thất bại. Lỗi: {e}")
            if temp_xlw_book: temp_xlw_book.close()
            return None

    def quit(self):
        """Đóng ứng dụng Excel."""
        if self._app:
            self._app.quit()
            self._app = None

    # --- Static Methods for Process Management ---
    @staticmethod
    def kill_all_processes():
        """(Phương thức tĩnh) Buộc đóng TẤT CẢ các tiến trình 'excel.exe'."""
        print("WARNING: Đang thực hiện buộc đóng TẤT CẢ các tiến trình Excel...")
        try:
            os.system('taskkill /F /IM excel.exe')
        except Exception as e:
            print(f"ERROR: Không thể thực thi lệnh taskkill. Lỗi: {e}")

    @staticmethod
    def kill_hidden_processes():
        """
        (Phương thức tĩnh) Chỉ tìm và buộc đóng các tiến trình Excel đang chạy ẩn (headless).
        An toàn hơn kill_all_processes vì nó không ảnh hưởng đến các file Excel người dùng đang mở.
        """
        print("INFO: Đang tìm và đóng các tiến trình Excel chạy ẩn...")
        killed_pids = []
        try:
            # Lấy danh sách tất cả các app đang chạy mà xlwings có thể thấy
            running_apps = xw.apps
            for app in running_apps:
                if not app.visible:
                    pid = app.pid
                    print(f"INFO: Tìm thấy tiến trình ẩn với PID: {pid}. Đang đóng...")
                    try:
                        app.quit() # Thử đóng nhẹ nhàng trước
                    except Exception:
                        # Nếu không được, dùng lệnh mạnh
                        os.system(f'taskkill /F /PID {pid}')
                    killed_pids.append(pid)
            
            if not killed_pids:
                print("INFO: Không tìm thấy tiến trình Excel nào đang chạy ẩn.")
            else:
                print(f"SUCCESS: Đã đóng thành công các tiến trình ẩn có PID: {killed_pids}")

        except Exception as e:
            print(f"ERROR: Đã xảy ra lỗi khi tìm và đóng tiến trình ẩn. Lỗi: {e}")
