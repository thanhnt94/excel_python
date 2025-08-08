# excel_toolkit/excelapp.py

import xlwings as xw
import os
import time
from pathlib import Path
from .workbook import Workbook  # Sử dụng import tương đối

class ExcelApp:
    """
    Lớp quản lý chính, đại diện cho một tiến trình (instance) của ứng dụng Excel.
    """

    def __init__(self, visible=True, add_book=False):
        """
        Khởi tạo ứng dụng Excel.
        """
        print("INFO: Khởi tạo tiến trình Excel...")
        try:
            self._app = xw.App(visible=visible, add_book=add_book)
        except Exception as e:
            print(f"ERROR: Không thể khởi tạo Excel. Vui lòng kiểm tra cài đặt của bạn. Lỗi: {e}")
            raise

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.quit()

    def open(self, path, password=None, read_only=False):
        """
        Mở một workbook. Nếu workbook đã được mở, nó sẽ trả về đối tượng đó thay vì mở lại.
        """
        file_path = Path(path).resolve()

        for book in self._app.books:
            if Path(book.fullname).resolve() == file_path:
                print(f"INFO: Gắn vào workbook đã mở: {file_path.name}")
                return Workbook(book, self)

        print(f"INFO: Đang mở workbook từ: {file_path}")
        try:
            xlw_book = self._app.books.open(
                file_path,
                password=password,
                read_only=read_only,
                ignore_read_only_recommended=True
            )
            return Workbook(xlw_book, self)
        except Exception as e:
            print(f"ERROR: Không thể mở workbook tại '{file_path}'. Lỗi: {e}")
            return None

    def new(self):
        """
        Tạo một workbook mới (trắng).
        """
        print("INFO: Đang tạo một workbook mới...")
        xlw_book = self._app.books.add()
        return Workbook(xlw_book, self)

    def get_workbook(self, specifier=None):
        """
        Lấy một workbook đã mở dựa trên một tiêu chí cụ thể (tên, index, hoặc active).
        """
        if not self._app.books:
            print("INFO: Không có workbook nào đang mở.")
            return None
        
        # Mặc định lấy workbook đang active
        if specifier is None:
            return self.get_active_workbook()

        try:
            # Thử lấy bằng tên hoặc index
            xlw_book = self._app.books[specifier]
            print(f"INFO: Gắn vào workbook theo tiêu chí '{specifier}'. Tên file: {xlw_book.name}")
            return Workbook(xlw_book, self)
        except Exception:
            print(f"ERROR: Không tìm thấy workbook nào khớp với '{specifier}'.")
            return None

    def get_active_workbook(self):
        """
        Lấy workbook đang được kích hoạt (active) trong ứng dụng Excel.
        """
        if self._app.books.active:
            active_book = self._app.books.active
            print(f"INFO: Gắn vào workbook đang active: {active_book.name}")
            return Workbook(active_book, self)
        print("INFO: Không có workbook nào đang được active.")
        return None

    def wait_for_workbook(self, title_contains=None, title_is=None, timeout=30):
        """
        Chờ cho đến khi một workbook thỏa mãn điều kiện xuất hiện.

        Args:
            title_contains (str, optional): Chờ workbook có tên chứa chuỗi này.
            title_is (str, optional): Chờ workbook có tên chính xác là chuỗi này.
            timeout (int, optional): Thời gian chờ tối đa (giây). Mặc định là 30.

        Returns:
            Workbook: Đối tượng Workbook tìm thấy, hoặc None nếu hết thời gian chờ.
        """
        print(f"INFO: Đang chờ workbook (timeout={timeout}s)... Điều kiện: "
              f"{'tên chứa ' + title_contains if title_contains else ''}"
              f"{'tên là ' + title_is if title_is else ''}")

        start_time = time.time()
        while time.time() - start_time < timeout:
            for book in self._app.books:
                match = False
                if title_is and book.name == title_is:
                    match = True
                elif title_contains and title_contains in book.name:
                    match = True
                
                if match:
                    print(f"SUCCESS: Đã tìm thấy workbook '{book.name}'.")
                    return Workbook(book, self)
            
            time.sleep(1) # Chờ 1 giây trước khi kiểm tra lại

        print(f"ERROR: Hết thời gian chờ ({timeout}s). Không tìm thấy workbook thỏa mãn điều kiện.")
        return None

    def quit(self):
        """
        Đóng ứng dụng Excel và tất cả các workbook liên quan.
        """
        if self._app:
            print("INFO: Đang đóng tiến trình Excel...")
            for book in list(self._app.books):
                book.close()
            self._app.quit()
            self._app = None
            print("INFO: Tiến trình Excel đã được đóng.")

    @staticmethod
    def kill_rogue_processes():
        """
        (Phương thức tĩnh) Buộc đóng tất cả các tiến trình 'excel.exe' đang chạy trên hệ thống.
        """
        print("WARNING: Đang thực hiện buộc đóng tất cả các tiến trình Excel...")
        try:
            os.system('taskkill /F /IM excel.exe')
            print("INFO: Đã gửi lệnh buộc đóng.")
        except Exception as e:
            print(f"ERROR: Không thể thực thi lệnh taskkill. Lỗi: {e}")

