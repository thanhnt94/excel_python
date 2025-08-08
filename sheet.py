# -*- coding: utf-8 -*-
"""
File: sheet.py
Author: Your Name / Tên của bạn
Description: Chứa class Sheet để đại diện và thao tác với một trang tính Excel.

--- CHANGELOG ---
Version 0.2.0 (2025-08-08):
    - Tích hợp các class Range và Shape.
    - Thêm các phương thức "nhà máy" để tạo và quản lý Range và Shape.
    - Methods added: .range(), .cell(), .shapes, .get_shape(), .add_textbox(), .find().

Version 0.1.0 (2025-08-08):
    - Khởi tạo class Sheet với các thuộc tính và phương thức cơ bản.
    - Properties: .name, .workbook, .is_visible.
    - State Management: .activate(), .hide(), .unhide(), .copy().
    - Boundary Detection: .get_last_row(), .get_last_column().
    - Content Management: .clear(), .clear_contents(), .autofit().
-------------------
"""

# Bước 1: Import các class đã tạo
from .range import Range
from .shape import Shape

class Sheet:
    """
    Đại diện cho một trang tính (worksheet).
    Đây là nơi thực hiện hầu hết các thao tác với dữ liệu, ô và các đối tượng.
    """
    def __init__(self, xlw_sheet, workbook_instance):
        self._xlw_sheet = xlw_sheet
        self._workbook = workbook_instance

    def __repr__(self):
        return f"<Sheet [{self.name}] in [{self.workbook.name}]>"

    # --- Properties ---
    @property
    def name(self):
        return self._xlw_sheet.name
    
    @name.setter
    def name(self, new_name):
        self._xlw_sheet.name = new_name

    @property
    def workbook(self):
        return self._workbook
        
    @property
    def is_visible(self):
        return self._xlw_sheet.api.Visible == -1

    # --- State Management ---
    def activate(self):
        self._xlw_sheet.activate()
        return self

    def hide(self):
        self._xlw_sheet.api.Visible = 0
        return self

    def unhide(self):
        self._xlw_sheet.api.Visible = -1
        return self

    def copy(self, new_name=None, before=None, after=None):
        before_sheet = self.workbook._xlw_book.sheets[before.name if isinstance(before, Sheet) else before] if before else None
        after_sheet = self.workbook._xlw_book.sheets[after.name if isinstance(after, Sheet) else after] if after else None
        self._xlw_sheet.copy(before=before_sheet, after=after_sheet)
        new_xlw_sheet = self.workbook._xlw_book.sheets.active
        if new_name:
            new_xlw_sheet.name = new_name
        return Sheet(new_xlw_sheet, self.workbook)

    # --- Navigation & Boundary ---
    def get_last_row(self, column=1):
        return self._xlw_sheet.range((self._xlw_sheet.cells.rows.count, column)).end('up').row

    def get_last_column(self, row=1):
        return self._xlw_sheet.range((row, self._xlw_sheet.cells.columns.count)).end('left').column

    # --- Data & Content ---
    def clear(self):
        self._xlw_sheet.clear()
        return self

    def clear_contents(self):
        self._xlw_sheet.used_range.clear_contents()
        return self

    def autofit(self, axis='columns'):
        if axis.lower().startswith('col'):
            self._xlw_sheet.autofit('c')
        elif axis.lower().startswith('row'):
            self._xlw_sheet.autofit('r')
        return self

    # --- Bước 2: Tích hợp Range (Factory Methods) ---
    def range(self, address):
        """
        Lấy một đối tượng Range từ một địa chỉ.

        Args:
            address (str): Địa chỉ của vùng (ví dụ: 'A1', 'A1:D10').

        Returns:
            Range: Một đối tượng Range của chúng ta.
        """
        xlw_range = self._xlw_sheet.range(address)
        return Range(xlw_range, self)

    def cell(self, row, column):
        """
        Lấy một đối tượng Range cho một ô duy nhất.

        Args:
            row (int): Số thứ tự dòng.
            column (int): Số thứ tự cột.

        Returns:
            Range: Một đối tượng Range của chúng ta.
        """
        xlw_range = self._xlw_sheet.range((row, column))
        return Range(xlw_range, self)

    def find(self, text, after='A1'):
        """
        Tìm ô đầu tiên chứa một giá trị cụ thể.

        Args:
            text (str): Nội dung cần tìm.
            after (str, optional): Địa chỉ ô bắt đầu tìm kiếm. Mặc định là 'A1'.

        Returns:
            Range: Đối tượng Range của ô tìm thấy, hoặc None.
        """
        found_cell = self._xlw_sheet.api.UsedRange.Find(What=text, After=self._xlw_sheet.api.Range(after))
        if found_cell:
            return self.cell(found_cell.Row, found_cell.Column)
        return None

    # --- Bước 3: Tích hợp Shape (Factory Methods) ---
    @property
    def shapes(self):
        """Trả về một danh sách tất cả các đối tượng Shape trong sheet."""
        return [Shape(s, self) for s in self._xlw_sheet.shapes]

    def get_shape(self, name_or_index):
        """Lấy một shape cụ thể bằng tên hoặc index."""
        try:
            xlw_shape = self._xlw_sheet.shapes[name_or_index]
            return Shape(xlw_shape, self)
        except Exception:
            print(f"ERROR: Không tìm thấy shape '{name_or_index}'.")
            return None

    def add_textbox(self, text, left, top, width, height):
        """
        Thêm một textbox mới vào sheet.

        Returns:
            Shape: Đối tượng Shape của textbox vừa được tạo.
        """
        print(f"INFO: Đang thêm textbox vào sheet '{self.name}'...")
        xlw_shape = self._xlw_sheet.shapes.add_textbox(left, top, width, height)
        new_shape = Shape(xlw_shape, self)
        new_shape.text = text
        return new_shape
