# -*- coding: utf-8 -*-
"""
File: workbook.py
Author: Your Name / Tên của bạn
Description: Chứa class Workbook để đại diện và quản lý một file Excel.

--- CHANGELOG ---
Version 0.3.0 (2025-08-08):
    - Nâng cấp .delete_sheet() và .delete_hidden_sheets() với tùy chọn 'safe=True'
      để phá vỡ các liên kết công thức trước khi xóa, tránh lỗi #REF!.

Version 0.2.0 (2025-08-08):
    - Thêm các phương thức nâng cao: .for_each_sheet(), quản lý Named Range, bảo vệ, tính toán.

Version 0.1.0 (2025-08-08):
    - Khởi tạo class Workbook với các chức năng cơ bản.
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

    # --- File Lifecycle & Calculation ---
    def save(self):
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
        new_xlw_sheet = self._xlw_book.sheets.add(name, before=self._xlw_book.sheets[before] if before else None, after=self._xlw_book.sheets[after] if after else None)
        return Sheet(new_xlw_sheet, self)

    def delete_sheet(self, specifier, safe=False):
        """
        Xóa một sheet theo tên hoặc index.

        Args:
            specifier (str or int): Tên hoặc index của sheet cần xóa.
            safe (bool, optional): Nếu True, sẽ tìm và phá vỡ các công thức
                                   tham chiếu đến sheet này trước khi xóa.
                                   Mặc định là False.
        """
        try:
            sheet_to_delete = self._xlw_book.sheets[specifier]
            sheet_name_to_delete = sheet_to_delete.name

            if safe:
                self._break_links_to_sheet(sheet_name_to_delete)

            self.app._app.display_alerts = False
            sheet_to_delete.delete()
            self.app._app.display_alerts = True
            print(f"SUCCESS: Đã xóa thành công sheet '{sheet_name_to_delete}'.")
        except Exception as e:
            print(f"ERROR: Không thể xóa sheet '{specifier}'. Lỗi: {e}")
            self.app._app.display_alerts = True
        return self

    def delete_hidden_sheets(self, safe=False):
        """
        Xóa tất cả các sheet đang bị ẩn.

        Args:
            safe (bool, optional): Nếu True, sẽ phá vỡ các công thức tham chiếu
                                   đến các sheet này trước khi xóa. Mặc định là False.
        """
        hidden_names = [sheet.name for sheet in self.hidden_sheets]
        if not hidden_names:
            print("INFO: Không có sheet ẩn nào để xóa.")
            return self

        print(f"INFO: Chuẩn bị xóa {len(hidden_names)} sheet ẩn. Chế độ an toàn: {safe}.")

        if safe:
            for name in hidden_names:
                self._break_links_to_sheet(name)
        
        for name in hidden_names:
            # Gọi hàm delete_sheet đã có sẵn (không cần safe nữa vì đã xử lý)
            self.delete_sheet(name, safe=False)
        
        print(f"SUCCESS: Đã hoàn tất việc xóa các sheet ẩn.")
        return self
        
    def _break_links_to_sheet(self, sheet_name_to_delete):
        """
        (Hàm nội bộ) Duyệt qua các sheet còn lại và thay thế các công thức
        tham chiếu đến 'sheet_name_to_delete' bằng giá trị tĩnh của chúng.
        """
        print(f"INFO (Safe Mode): Đang tìm và phá vỡ các liên kết đến sheet '{sheet_name_to_delete}'...")
        sheets_to_keep = [s for s in self.sheets if s.name != sheet_name_to_delete]
        
        for sheet in sheets_to_keep:
            # Chỉ tìm trong các ô đã được sử dụng để tăng tốc
            for cell in sheet.used_range._xlw_range:
                if cell.has_formula:
                    # Kiểm tra một cách đơn giản nhưng hiệu quả
                    if f"'{sheet_name_to_delete}'!" in cell.formula or f"{sheet_name_to_delete}!" in cell.formula:
                        print(f"  - Tìm thấy tham chiếu tại {sheet.name}!{cell.address}. Đang thay thế bằng giá trị...")
                        # Tuyệt chiêu của xlwings: đọc giá trị và ghi đè lại
                        cell.value = cell.value

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
