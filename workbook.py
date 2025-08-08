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
        Lớp này không nên được người dùng tạo trực tiếp, mà thông qua ExcelApp.open() hoặc ExcelApp.new().

        Args:
            xlw_book (xlwings.Book): Đối tượng Book gốc từ thư viện xlwings.
            app_instance (ExcelApp): Đối tượng ExcelApp đã tạo ra workbook này.
        """
        self._xlw_book = xlw_book
        self._app = app_instance

    def __repr__(self):
        """Biểu diễn đối tượng dưới dạng chuỗi, hữu ích cho việc debug."""
        return f"<Workbook [{self.name}]>"

    # --- Properties ---
    @property
    def name(self):
        """Trả về tên của file workbook (ví dụ: 'report.xlsx')."""
        return self._xlw_book.name

    @property
    def path(self):
        """Trả về đường dẫn đầy đủ của file dưới dạng đối tượng Path."""
        return Path(self._xlw_book.fullname)

    @property
    def app(self):
        """Trả về đối tượng ExcelApp cha."""
        return self._app

    @property
    def sheets(self):
        """Trả về một danh sách tất cả các đối tượng Sheet có trong workbook."""
        # Bọc mỗi sheet của xlwings trong class Sheet của chúng ta
        return [Sheet(s, self) for s in self._xlw_book.sheets]

    @property
    def visible_sheets(self):
        """Trả về một danh sách chỉ các sheet đang được hiển thị."""
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible == -1] # -1 is xlSheetVisible

    @property
    def hidden_sheets(self):
        """Trả về một danh sách chỉ các sheet đang bị ẩn (bao gồm cả rất ẩn)."""
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible != -1]

    @property
    def sheet_names(self):
        """Trả về một list tên của tất cả các sheet."""
        return [s.name for s in self._xlw_book.sheets]

    # --- File Lifecycle Management ---
    def save(self):
        """
        Lưu các thay đổi vào file hiện tại.

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        print(f"INFO: Đang lưu workbook '{self.name}'...")
        self._xlw_book.save()
        return self

    def save_as(self, new_path):
        """
        Lưu workbook với một tên mới hoặc ở một vị trí khác.

        Args:
            new_path (str or Path): Đường dẫn file mới.

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        print(f"INFO: Đang lưu workbook thành '{new_path}'...")
        self._xlw_book.save(new_path)
        return self

    def close(self, save_changes=False):
        """
        Đóng workbook.

        Args:
            save_changes (bool, optional): True để lưu các thay đổi trước khi đóng.
                                           Mặc định là False.
        """
        print(f"INFO: Đang đóng workbook '{self.name}'...")
        if save_changes:
            self.save()
        self._xlw_book.close()
        # Sau khi đóng, đối tượng này không còn hợp lệ, không return self

    def activate(self):
        """
        Kích hoạt (đưa lên phía trước) workbook này.

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        print(f"INFO: Đang kích hoạt workbook '{self.name}'...")
        self._xlw_book.activate()
        return self

    # --- Sheet Management ---
    def sheet(self, specifier):
        """
        Lấy một sheet cụ thể bằng tên hoặc index.

        Args:
            specifier (str or int): Tên (ví dụ: 'Sheet1') hoặc index (ví dụ: 0) của sheet.

        Returns:
            Sheet: Đối tượng Sheet tìm thấy, hoặc None nếu không có.
        """
        try:
            xlw_sheet = self._xlw_book.sheets[specifier]
            return Sheet(xlw_sheet, self)
        except Exception:
            print(f"ERROR: Không tìm thấy sheet '{specifier}' trong workbook '{self.name}'.")
            return None
    
    def add_sheet(self, name, before=None, after=None):
        """
        Thêm một sheet mới.

        Args:
            name (str): Tên của sheet mới.
            before (Sheet or str, optional): Sheet hoặc tên sheet để chèn vào trước.
            after (Sheet or str, optional): Sheet hoặc tên sheet để chèn vào sau.

        Returns:
            Sheet: Đối tượng Sheet vừa được tạo.
        """
        print(f"INFO: Đang thêm sheet '{name}'...")
        before_sheet = self._xlw_book.sheets[before] if before else None
        after_sheet = self._xlw_book.sheets[after] if after else None
        
        new_xlw_sheet = self._xlw_book.sheets.add(name, before=before_sheet, after=after_sheet)
        return Sheet(new_xlw_sheet, self)

    # --- Conversion & Publishing ---
    def to_pdf(self, output_path=None, quality='standard'):
        """
        Chuyển đổi toàn bộ workbook thành file PDF.

        Args:
            output_path (str or Path, optional): Đường dẫn file PDF output. 
                                                 Nếu None, sẽ lưu cùng thư mục với file Excel.
            quality (str, optional): Chất lượng ('standard' hoặc 'minimum').

        Returns:
            Workbook: Trả về chính nó để cho phép nối chuỗi phương thức.
        """
        if not output_path:
            output_path = self.path.with_suffix('.pdf')
        else:
            output_path = Path(output_path)

        print(f"INFO: Đang chuyển đổi '{self.name}' sang PDF tại '{output_path}'...")
        self.activate() # Đảm bảo workbook được kích hoạt trước khi in
        time.sleep(1) # Chờ một chút để Excel xử lý
        
        # 0 = xlQualityStandard, 1 = xlQualityMinimum
        quality_val = 0 if quality == 'standard' else 1
        self._xlw_book.api.ExportAsFixedFormat(0, str(output_path), Quality=quality_val)
        print("INFO: Chuyển đổi PDF thành công.")
        return self

