# -*- coding: utf-8 -*-
"""
File: range.py
Author: Your Name / Tên của bạn
Description: Chứa class Range để đại diện và thao tác với một ô hoặc một vùng ô.

--- CHANGELOG ---
Version 0.1.0 (2025-08-08):
    - Khởi tạo class Range.
    - Properties: .value, .formula, .address, .sheet, .row, .column.
    - Methods: .clear(), .clear_contents(), .copy_to(), .style(), .merge(), .unmerge(), .autofit().
-------------------
"""

class Range:
    """
    Đại diện cho một ô hoặc một vùng ô trong một sheet.
    Cung cấp các phương thức để đọc, ghi và định dạng dữ liệu.
    """
    def __init__(self, xlw_range, sheet_instance):
        self._xlw_range = xlw_range
        self._sheet = sheet_instance

    def __repr__(self):
        return f"<Range [{self.address}] on Sheet [{self.sheet.name}]>"

    # --- Properties ---
    @property
    def value(self):
        """Lấy hoặc đặt giá trị cho vùng."""
        return self._xlw_range.value
    
    @value.setter
    def value(self, data):
        self._xlw_range.value = data

    @property
    def formula(self):
        """Lấy hoặc đặt công thức cho vùng."""
        return self._xlw_range.formula

    @formula.setter
    def formula(self, formula_string):
        self._xlw_range.formula = formula_string

    @property
    def address(self):
        """Trả về địa chỉ của vùng (ví dụ: '$A$1:$B$10')."""
        return self._xlw_range.address

    @property
    def sheet(self):
        """Trả về đối tượng Sheet cha."""
        return self._sheet

    @property
    def row(self):
        """Trả về số thứ tự dòng bắt đầu của vùng."""
        return self._xlw_range.row

    @property
    def column(self):
        """Trả về số thứ tự cột bắt đầu của vùng."""
        return self._xlw_range.column

    # --- Content Management ---
    def clear(self):
        """Xóa tất cả nội dung và định dạng của vùng."""
        self._xlw_range.clear()
        return self

    def clear_contents(self):
        """Chỉ xóa nội dung, giữ lại định dạng."""
        self._xlw_range.clear_contents()
        return self

    def copy_to(self, destination):
        """
        Sao chép vùng này đến một vị trí mới.

        Args:
            destination (Range or str): Đối tượng Range hoặc địa chỉ ô đích (ví dụ: 'D1').
        """
        dest_range = destination._xlw_range if isinstance(destination, Range) else self.sheet._xlw_sheet.range(destination)
        self._xlw_range.copy(dest_range)
        return self

    # --- Formatting ---
    def style(self, font_bold=None, font_italic=None, font_color=None, interior_color=None, number_format=None):
        """
        Áp dụng nhiều kiểu định dạng cho vùng trong một lần gọi.

        Args:
            font_bold (bool, optional): In đậm.
            font_italic (bool, optional): In nghiêng.
            font_color (str or tuple, optional): Màu chữ (ví dụ: '#FF0000' hoặc (255, 0, 0)).
            interior_color (str or tuple, optional): Màu nền.
            number_format (str, optional): Định dạng số (ví dụ: '0.00%', '#,##0').
        """
        if font_bold is not None:
            self._xlw_range.font.bold = font_bold
        if font_italic is not None:
            self._xlw_range.font.italic = font_italic
        if font_color:
            self._xlw_range.font.color = font_color
        if interior_color:
            self._xlw_range.color = interior_color
        if number_format:
            self._xlw_range.number_format = number_format
        return self

    def merge(self):
        """Hợp nhất các ô trong vùng này thành một ô duy nhất."""
        self._xlw_range.merge()
        return self

    def unmerge(self):
        """Tách các ô đã được hợp nhất."""
        self._xlw_range.unmerge()
        return self

    def autofit(self):
        """Tự động điều chỉnh độ rộng cột và chiều cao hàng của vùng này."""
        self._xlw_range.autofit()
        return self
