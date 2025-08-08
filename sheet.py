# -*- coding: utf-8 -*-
"""
File: sheet.py
Author: Your Name / Tên của bạn
Description: Chứa class Sheet để đại diện và thao tác với một trang tính Excel.

--- CHANGELOG ---
Version 0.1.0 (2025-08-08):
    - Khởi tạo class Sheet với các thuộc tính và phương thức cơ bản.
    - Properties: .name, .workbook, .is_visible.
    - State Management: .activate(), .hide(), .unhide(), .copy().
    - Boundary Detection: .get_last_row(), .get_last_column().
    - Content Management: .clear(), .clear_contents(), .autofit().
-------------------
"""

class Sheet:
    """
    Đại diện cho một trang tính (worksheet).
    Đây là nơi thực hiện hầu hết các thao tác với dữ liệu, ô và các đối tượng.
    """
    def __init__(self, xlw_sheet, workbook_instance):
        """
        Khởi tạo một đối tượng Sheet.
        Không nên được tạo trực tiếp, mà thông qua workbook.sheet().

        Args:
            xlw_sheet (xlwings.Sheet): Đối tượng Sheet gốc từ xlwings.
            workbook_instance (Workbook): Đối tượng Workbook cha.
        """
        self._xlw_sheet = xlw_sheet
        self._workbook = workbook_instance

    def __repr__(self):
        return f"<Sheet [{self.name}] in [{self.workbook.name}]>"

    # --- Properties ---
    @property
    def name(self):
        """Lấy hoặc đặt tên cho sheet."""
        return self._xlw_sheet.name
    
    @name.setter
    def name(self, new_name):
        print(f"INFO: Đổi tên sheet từ '{self.name}' thành '{new_name}'...")
        self._xlw_sheet.name = new_name

    @property
    def workbook(self):
        """Trả về đối tượng Workbook cha."""
        return self._workbook
        
    @property
    def is_visible(self):
        """Kiểm tra xem sheet có đang hiển thị hay không (True/False)."""
        return self._xlw_sheet.api.Visible == -1

    # --- State Management ---
    def activate(self):
        """Kích hoạt (chọn) sheet này."""
        print(f"INFO: Đang kích hoạt sheet '{self.name}'...")
        self._xlw_sheet.activate()
        return self

    def hide(self):
        """Ẩn sheet này."""
        print(f"INFO: Đang ẩn sheet '{self.name}'...")
        self._xlw_sheet.api.Visible = 0 # 0 = xlSheetHidden
        return self

    def unhide(self):
        """Hiện lại sheet này."""
        print(f"INFO: Đang hiện lại sheet '{self.name}'...")
        self._xlw_sheet.api.Visible = -1 # -1 = xlSheetVisible
        return self

    def copy(self, new_name=None, before=None, after=None):
        """
        Tạo một bản sao của sheet này trong cùng một workbook.

        Args:
            new_name (str, optional): Tên cho sheet mới.
            before (Sheet or str, optional): Sheet hoặc tên sheet để chèn vào trước.
            after (Sheet or str, optional): Sheet hoặc tên sheet để chèn vào sau.

        Returns:
            Sheet: Đối tượng Sheet mới được tạo ra.
        """
        print(f"INFO: Đang sao chép sheet '{self.name}'...")
        # Lấy đối tượng sheet gốc từ tên hoặc đối tượng Sheet của chúng ta
        before_sheet = self.workbook._xlw_book.sheets[before.name if isinstance(before, Sheet) else before] if before else None
        after_sheet = self.workbook._xlw_book.sheets[after.name if isinstance(after, Sheet) else after] if after else None

        self._xlw_sheet.copy(before=before_sheet, after=after_sheet)
        
        # xlwings tự động kích hoạt sheet mới, ta lấy nó và bọc lại
        new_xlw_sheet = self.workbook._xlw_book.sheets.active
        if new_name:
            new_xlw_sheet.name = new_name
            
        print(f"SUCCESS: Đã tạo bản sao với tên '{new_xlw_sheet.name}'.")
        return Sheet(new_xlw_sheet, self.workbook)

    # --- Navigation & Boundary ---
    def get_last_row(self, column=1):
        """
        Tìm dòng cuối cùng có dữ liệu trong một cột cụ thể.

        Args:
            column (int or str, optional): Số thứ tự hoặc tên cột (ví dụ: 'A'). Mặc định là 1.

        Returns:
            int: Số thứ tự của dòng cuối cùng.
        """
        # Phương pháp .end('up') đáng tin cậy
        return self._xlw_sheet.range(
            (self._xlw_sheet.cells.rows.count, column)
        ).end('up').row

    def get_last_column(self, row=1):
        """
        Tìm cột cuối cùng có dữ liệu trong một hàng cụ thể.

        Args:
            row (int, optional): Số thứ tự hàng. Mặc định là 1.

        Returns:
            int: Số thứ tự của cột cuối cùng.
        """
        return self._xlw_sheet.range(
            (row, self._xlw_sheet.cells.columns.count)
        ).end('left').column

    # --- Data & Content ---
    def clear(self):
        """Xóa tất cả nội dung và định dạng khỏi sheet."""
        print(f"WARNING: Đang xóa toàn bộ nội dung và định dạng của sheet '{self.name}'...")
        self._xlw_sheet.clear()
        return self

    def clear_contents(self):
        """Chỉ xóa nội dung, giữ lại định dạng."""
        print(f"INFO: Đang xóa nội dung của sheet '{self.name}'...")
        self._xlw_sheet.used_range.clear_contents()
        return self

    def autofit(self, axis='columns'):
        """
        Tự động điều chỉnh độ rộng cột hoặc chiều cao hàng để vừa với nội dung.

        Args:
            axis (str, optional): 'columns' hoặc 'rows'. Mặc định là 'columns'.
        """
        if axis.lower().startswith('col'):
            print(f"INFO: Tự động điều chỉnh cột cho sheet '{self.name}'...")
            self._xlw_sheet.autofit('c')
        elif axis.lower().startswith('row'):
            print(f"INFO: Tự động điều chỉnh hàng cho sheet '{self.name}'...")
            self._xlw_sheet.autofit('r')
        return self

    # --- Các phương thức cho Range, Find, Shape sẽ được thêm vào sau ---
    # def range(self, address): ...
    # def find(self, text): ...
    # def add_shape(self, ...): ...

