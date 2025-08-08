# excel_toolkit/workbook.py

from pathlib import Path
import time
from .sheet import Sheet # Import tương đối

class Workbook:
    """
    Đại diện cho một file Excel (sổ làm việc).
    Cung cấp các phương thức để quản lý file và các sheet bên trong nó.
    """

    def __init__(self, xlw_book, app_instance):
        """
        Khởi tạo một đối tượng Workbook.
        """
        self._xlw_book = xlw_book
        self._app = app_instance

    def __repr__(self):
        """Biểu diễn đối tượng dưới dạng chuỗi, hữu ích cho việc debug."""
        return f"<Workbook [{self.name}]>"

    # --- Properties ---
    @property
    def name(self):
        """Trả về tên của file workbook."""
        return self._xlw_book.name

    @property
    def path(self):
        """Trả về đường dẫn đầy đủ của file."""
        return Path(self._xlw_book.fullname)

    @property
    def app(self):
        """Trả về đối tượng ExcelApp cha."""
        return self._app

    @property
    def sheets(self):
        """Trả về một danh sách tất cả các đối tượng Sheet."""
        return [Sheet(s, self) for s in self._xlw_book.sheets]

    @property
    def visible_sheets(self):
        """Trả về một danh sách chỉ các sheet đang được hiển thị."""
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible == -1]

    @property
    def hidden_sheets(self):
        """Trả về một danh sách chỉ các sheet đang bị ẩn."""
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible != -1]

    @property
    def sheet_names(self):
        """Trả về một list tên của tất cả các sheet."""
        return [s.name for s in self._xlw_book.sheets]

    # --- File Lifecycle Management ---
    def save(self):
        """Lưu các thay đổi vào file hiện tại."""
        print(f"INFO: Đang lưu workbook '{self.name}'...")
        self._xlw_book.save()
        return self

    def save_as(self, new_path):
        """Lưu workbook với một tên mới."""
        print(f"INFO: Đang lưu workbook thành '{new_path}'...")
        self._xlw_book.save(new_path)
        return self

    def close(self, save_changes=False):
        """Đóng workbook."""
        print(f"INFO: Đang đóng workbook '{self.name}'...")
        if save_changes:
            self.save()
        self._xlw_book.close()

    def activate(self):
        """Kích hoạt (đưa lên phía trước) workbook này."""
        print(f"INFO: Đang kích hoạt workbook '{self.name}'...")
        self._xlw_book.activate()
        return self

    # --- Sheet Management ---
    def sheet(self, specifier):
        """Lấy một sheet cụ thể bằng tên hoặc index."""
        try:
            xlw_sheet = self._xlw_book.sheets[specifier]
            return Sheet(xlw_sheet, self)
        except Exception:
            print(f"ERROR: Không tìm thấy sheet '{specifier}' trong workbook '{self.name}'.")
            return None
    
    def add_sheet(self, name, before=None, after=None):
        """Thêm một sheet mới."""
        print(f"INFO: Đang thêm sheet '{name}'...")
        before_sheet = self._xlw_book.sheets[before] if before else None
        after_sheet = self._xlw_book.sheets[after] if after else None
        
        new_xlw_sheet = self._xlw_book.sheets.add(name, before=before_sheet, after=after_sheet)
        return Sheet(new_xlw_sheet, self)

    def delete_sheet(self, specifier):
        """
        Xóa một sheet (kể cả sheet ẩn) theo tên hoặc index.
        Hành động này không thể hoàn tác.

        Args:
            specifier (str or int): Tên hoặc index của sheet cần xóa.

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        print(f"WARNING: Đang chuẩn bị xóa sheet '{specifier}'...")
        try:
            sheet_to_delete = self._xlw_book.sheets[specifier]
            
            self.app._app.display_alerts = False
            sheet_to_delete.delete()
            self.app._app.display_alerts = True
            
            print(f"SUCCESS: Đã xóa thành công sheet '{specifier}'.")
        except Exception as e:
            print(f"ERROR: Không thể xóa sheet '{specifier}'. Lỗi: {e}")
            self.app._app.display_alerts = True
        
        return self

    def delete_hidden_sheets(self):
        """
        Xóa tất cả các sheet đang bị ẩn trong workbook.
        Hành động này không thể hoàn tác.

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        print("INFO: Đang tìm và xóa tất cả các sheet ẩn...")
        hidden_names = [sheet.name for sheet in self.hidden_sheets]
        
        if not hidden_names:
            print("INFO: Không có sheet ẩn nào để xóa.")
            return self

        for name in hidden_names:
            # Gọi hàm delete_sheet đã được định nghĩa ở trên
            self.delete_sheet(name)
        
        print(f"SUCCESS: Đã hoàn tất việc xóa các sheet ẩn.")
        return self

    # --- Conversion & Publishing ---
    def to_pdf(self, output_path=None, quality='standard'):
        """Chuyển đổi toàn bộ workbook thành file PDF."""
        if not output_path:
            output_path = self.path.with_suffix('.pdf')
        else:
            output_path = Path(output_path)

        print(f"INFO: Đang chuyển đổi '{self.name}' sang PDF tại '{output_path}'...")
        self.activate()
        time.sleep(1)
        
        quality_val = 0 if quality == 'standard' else 1
        self._xlw_book.api.ExportAsFixedFormat(0, str(output_path), Quality=quality_val)
        print("INFO: Chuyển đổi PDF thành công.")
        return self
