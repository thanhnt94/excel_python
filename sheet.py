# -*- coding: utf-8 -*-
"""
File: workbook.py
Author: Your Name / Tên của bạn
Description: Chứa class Workbook để đại diện và quản lý một file Excel.

--- CHANGELOG ---
Version 0.2.0 (2025-08-08):
    - Thêm phương thức .for_each_sheet() để thực thi hành động trên hàng loạt sheet.
    - Thêm các phương thức quản lý Named Range: .add_named_range(), .get_named_range(), .delete_named_range().
    - Thêm các phương thức bảo vệ: .protect(), .unprotect().
    - Thêm phương thức tính toán lại: .calculate().

Version 0.1.0 (2025-08-08):
    - Khởi tạo class Workbook với các chức năng cơ bản.
    - Methods: .save(), .save_as(), .close(), .activate(), .sheet(), .add_sheet(), .delete_sheet(), .to_pdf().
-------------------
"""

from pathlib import Path
import time
from .sheet import Sheet
from .range import Range

class Workbook:
    """
    Đại diện cho một file Excel (sổ làm việc).
    """
    def __init__(self, xlw_book, app_instance):
        self._xlw_book = xlw_book
        self._app = app_instance

    def __repr__(self):
        return f"<Workbook [{self.name}]>"

    # --- Properties ---
    @property
    def name(self):
        return self._xlw_book.name

    @property
    def path(self):
        return Path(self._xlw_book.fullname)

    @property
    def app(self):
        return self._app

    @property
    def sheets(self):
        return [Sheet(s, self) for s in self._xlw_book.sheets]

    @property
    def visible_sheets(self):
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible == -1]

    @property
    def hidden_sheets(self):
        return [Sheet(s, self) for s in self._xlw_book.sheets if s.api.Visible != -1]

    @property
    def sheet_names(self):
        return [s.name for s in self._xlw_book.sheets]

    # --- File Lifecycle & Calculation ---
    def save(self):
        self._xlw_book.save()
        return self

    def save_as(self, new_path):
        self._xlw_book.save(new_path)
        return self

    def close(self, save_changes=False):
        if save_changes:
            self.save()
        self._xlw_book.close()

    def activate(self):
        self._xlw_book.activate()
        return self

    def calculate(self):
        """Buộc Excel tính toán lại tất cả các công thức trong workbook."""
        print(f"INFO: Đang tính toán lại công thức cho workbook '{self.name}'...")
        self._xlw_book.api.Calculate()
        return self

    # --- Protection ---
    def protect(self, password=None):
        """Bảo vệ cấu trúc của workbook (ngăn thêm, xóa, di chuyển sheet)."""
        print(f"INFO: Đang bảo vệ workbook '{self.name}'...")
        self._xlw_book.api.Protect(Password=password)
        return self

    def unprotect(self, password=None):
        """Mở khóa bảo vệ cấu trúc của workbook."""
        print(f"INFO: Đang mở khóa bảo vệ cho workbook '{self.name}'...")
        self._xlw_book.api.Unprotect(Password=password)
        return self

    # --- Sheet Management ---
    def sheet(self, specifier):
        try:
            return Sheet(self._xlw_book.sheets[specifier], self)
        except Exception:
            return None
    
    def add_sheet(self, name, before=None, after=None):
        before_sheet = self._xlw_book.sheets[before] if before else None
        after_sheet = self._xlw_book.sheets[after] if after else None
        new_xlw_sheet = self._xlw_book.sheets.add(name, before=before_sheet, after=after_sheet)
        return Sheet(new_xlw_sheet, self)

    def delete_sheet(self, specifier):
        try:
            self.app._app.display_alerts = False
            self._xlw_book.sheets[specifier].delete()
            self.app._app.display_alerts = True
        except Exception as e:
            print(f"ERROR: Không thể xóa sheet '{specifier}'. Lỗi: {e}")
            self.app._app.display_alerts = True
        return self

    def delete_hidden_sheets(self):
        hidden_names = [sheet.name for sheet in self.hidden_sheets]
        if not hidden_names:
            return self
        for name in hidden_names:
            self.delete_sheet(name)
        return self

    def for_each_sheet(self, action, include=None, exclude=None):
        """
        Thực thi một hành động trên nhiều sheet.

        Args:
            action (function): Hàm để thực thi, nhận một đối tượng Sheet làm tham số.
            include (list, optional): Danh sách tên các sheet cần xử lý.
            exclude (list, optional): Danh sách tên các sheet cần bỏ qua.
        """
        target_sheets = self.sheets
        
        if include:
            target_sheets = [s for s in target_sheets if s.name in include]
        
        if exclude:
            target_sheets = [s for s in target_sheets if s.name not in exclude]

        print(f"INFO: Đang thực thi hành động trên {len(target_sheets)} sheet...")
        for sheet in target_sheets:
            try:
                action(sheet)
            except Exception as e:
                print(f"ERROR: Lỗi khi thực thi hành động trên sheet '{sheet.name}'. Lỗi: {e}")
        
        return self

    # --- Named Range Management ---
    def add_named_range(self, name, address):
        """Tạo một vùng được đặt tên (Named Range)."""
        print(f"INFO: Đang tạo Named Range '{name}' cho địa chỉ '{address}'...")
        self._xlw_book.names.add(name, f"='{self.sheets[0].name}'!{address}")
        return self

    def get_named_range(self, name):
        """Lấy một đối tượng Range từ một tên đã đặt."""
        try:
            xlw_range = self._xlw_book.range(name)
            # Tìm xem range đó thuộc sheet nào
            sheet_name = xlw_range.sheet.name
            sheet_obj = self.sheet(sheet_name)
            return Range(xlw_range, sheet_obj)
        except Exception:
            print(f"ERROR: Không tìm thấy Named Range với tên '{name}'.")
            return None

    def delete_named_range(self, name):
        """Xóa một vùng được đặt tên."""
        try:
            self._xlw_book.names[name].delete()
            print(f"SUCCESS: Đã xóa Named Range '{name}'.")
        except Exception:
            print(f"ERROR: Không tìm thấy Named Range để xóa với tên '{name}'.")
        return self

    # --- Conversion & Publishing ---
    def to_pdf(self, output_path=None, quality='standard'):
        if not output_path:
            output_path = self.path.with_suffix('.pdf')
        else:
            output_path = Path(output_path)
        
        self.activate()
        time.sleep(1)
        quality_val = 0 if quality == 'standard' else 1
        self._xlw_book.api.ExportAsFixedFormat(0, str(output_path), Quality=quality_val)
        return self
