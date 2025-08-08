# -*- coding: utf-8 -*-
"""
File: shape.py
Author: Your Name / Tên của bạn
Description: Chứa class Shape để đại diện và thao tác với các đối tượng đồ họa.

--- CHANGELOG ---
Version 0.1.0 (2025-08-08):
    - Khởi tạo class Shape.
    - Properties: .name, .text, .left, .top, .width, .height, .sheet.
    - Methods: .delete(), .copy().
-------------------
"""

class Shape:
    """
    Đại diện cho một đối tượng đồ họa (shape, textbox, picture) trong một sheet.
    """
    def __init__(self, xlw_shape, sheet_instance):
        self._xlw_shape = xlw_shape
        self._sheet = sheet_instance

    def __repr__(self):
        return f"<Shape [{self.name}] on Sheet [{self.sheet.name}]>"

    # --- Properties ---
    @property
    def name(self):
        """Lấy hoặc đặt tên cho shape."""
        return self._xlw_shape.name
    
    @name.setter
    def name(self, new_name):
        self._xlw_shape.name = new_name

    @property
    def text(self):
        """Lấy hoặc đặt nội dung văn bản của shape (nếu có)."""
        return self._xlw_shape.text
    
    @text.setter
    def text(self, new_text):
        self._xlw_shape.text = new_text

    @property
    def left(self):
        """Vị trí cạnh trái của shape."""
        return self._xlw_shape.left

    @left.setter
    def left(self, value):
        self._xlw_shape.left = value

    @property
    def top(self):
        """Vị trí cạnh trên của shape."""
        return self._xlw_shape.top

    @top.setter
    def top(self, value):
        self._xlw_shape.top = value

    @property
    def width(self):
        """Độ rộng của shape."""
        return self._xlw_shape.width

    @width.setter
    def width(self, value):
        self._xlw_shape.width = value

    @property
    def height(self):
        """Chiều cao của shape."""
        return self._xlw_shape.height

    @height.setter
    def height(self, value):
        self._xlw_shape.height = value

    @property
    def sheet(self):
        """Trả về đối tượng Sheet cha."""
        return self._sheet

    # --- Actions ---
    def delete(self):
        """Xóa shape này."""
        print(f"INFO: Đang xóa shape '{self.name}'...")
        self._xlw_shape.delete()
        # Sau khi xóa, đối tượng này không còn hợp lệ, không return self

    def copy(self):
        """
        Sao chép shape này vào clipboard.
        Lưu ý: Việc dán (paste) sẽ là một phương thức của Sheet.
        """
        print(f"INFO: Đang sao chép shape '{self.name}' vào clipboard...")
        self._xlw_shape.api.Copy()
        return self # Có thể return self để nối chuỗi nếu cần
