# excel_toolkit/excelapp.py

import xlwings as xw
import os
from pathlib import Path
from .workbook import Workbook  # Sử dụng import tương đối để gọi đến file workbook.py

class ExcelApp:
    """
    Lớp quản lý chính, đại diện cho một tiến trình (instance) của ứng dụng Excel.

    Hoạt động như một context manager để đảm bảo việc khởi tạo và
    đóng ứng dụng được xử lý một cách an toàn và tự động.
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

    def get_active_workbook(self):
        """
        Lấy workbook đang được kích hoạt (active) trong ứng dụng Excel.

        Returns:
            Workbook: Đối tượng Workbook đang active, hoặc None nếu không có.
        """
        if not self._app.books:
            print("INFO: Không có workbook nào đang mở.")
            return None
        
        active_book = self._app.books.active
        if active_book:
            print(f"INFO: Gắn vào workbook đang active: {active_book.name}")
            return Workbook(active_book, self)
        
        print("INFO: Không có workbook nào đang được active.")
        return None

    def get_workbook(self, specifier=None):
        """
        Lấy một workbook đã mở dựa trên một tiêu chí cụ thể.

        Args:
            specifier (str or int, optional):
                - Nếu là str: Tên của workbook (ví dụ: 'BaoCao.xlsx').
                - Nếu là int: Index của workbook trong danh sách (0 là file mở đầu tiên, -1 là file mở cuối cùng).
                - Nếu là None (mặc định): Sẽ lấy workbook đang active.

        Returns:
            Workbook: Đối tượng Workbook tìm thấy, hoặc None nếu không có.
        """
        if not self._app.books:
            print("INFO: Không có workbook nào đang mở.")
            return None

        if specifier is None:
            return self.get_active_workbook()

        if isinstance(specifier, str):
            try:
                xlw_book = self._app.books[specifier]
                print(f"INFO: Gắn vào workbook theo tên: {specifier}")
                return Workbook(xlw_book, self)
            except Exception:
                print(f"ERROR: Không tìm thấy workbook nào có tên '{specifier}'.")
                return None
        
        if isinstance(specifier, int):
            try:
                xlw_book = self._app.books[specifier]
                print(f"INFO: Gắn vào workbook theo index: {specifier} (Tên file: {xlw_book.name})")
                return Workbook(xlw_book, self)
            except IndexError:
                print(f"ERROR: Index {specifier} nằm ngoài phạm vi. Số workbook đang mở: {len(self._app.books)}.")
                return None
        
        print(f"ERROR: Loại specifier không hợp lệ: {type(specifier)}.")
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
