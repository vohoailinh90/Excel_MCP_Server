"""
ExcelBackend: interface trừu tượng.

Server MCP (server.py) chỉ gọi các method ở đây, không biết bên dưới
là win32com/xlwings (thật, chạy trên Windows + Excel) hay openpyxl
(mock, dùng để test logic mà không cần Excel thật).

Mọi backend PHẢI implement đủ các method này với cùng chữ ký & kiểu
trả về, để có thể swap qua lại mà không sửa server.py.
"""

from abc import ABC, abstractmethod
from typing import Any


class ExcelBackendError(Exception):
    """Lỗi nghiệp vụ (sai tên sheet, sai workbook_id, macro không tồn tại...)."""


class ExcelBackend(ABC):

    @abstractmethod
    def open_workbook(self, path: str, visible: bool = False) -> str:
        """Mở file Excel, trả về workbook_id (dùng cho các lệnh sau)."""

    @abstractmethod
    def close_workbook(self, workbook_id: str, save: bool = True) -> None:
        """Đóng workbook, lưu nếu save=True."""

    @abstractmethod
    def list_open_workbooks(self) -> list[dict[str, Any]]:
        """Danh sách workbook đang mở qua server (id, path)."""

    @abstractmethod
    def list_sheets(self, workbook_id: str) -> list[str]:
        """Danh sách tên sheet."""

    @abstractmethod
    def read_cell(self, workbook_id: str, sheet: str, cell: str) -> Any:
        """Đọc giá trị 1 ô, ví dụ cell='B5'."""

    @abstractmethod
    def read_range(self, workbook_id: str, sheet: str, cell_range: str) -> list[list[Any]]:
        """Đọc 1 vùng, ví dụ cell_range='A1:C10'. Trả về mảng 2 chiều."""

    @abstractmethod
    def write_cell(self, workbook_id: str, sheet: str, cell: str, value: Any) -> None:
        """Ghi giá trị vào 1 ô."""

    @abstractmethod
    def write_range(
        self, workbook_id: str, sheet: str, start_cell: str, values: list[list[Any]]
    ) -> None:
        """Ghi mảng 2 chiều bắt đầu từ start_cell."""

    @abstractmethod
    def apply_formula(self, workbook_id: str, sheet: str, cell: str, formula: str) -> None:
        """Gán công thức, ví dụ formula='=SUM(A1:A10)'."""

    @abstractmethod
    def create_sheet(
        self,
        workbook_id: str,
        name: str,
        copy_from: str | None = None,
        position: int | None = None,
    ) -> None:
        """Tạo sheet mới, có thể copy nội dung từ sheet khác (copy_from)."""

    @abstractmethod
    def run_macro(self, workbook_id: str, macro_name: str, args: list[Any] | None = None) -> Any:
        """Chạy macro theo tên, trả về giá trị macro return (nếu có)."""

    @abstractmethod
    def save_workbook(self, workbook_id: str, path: str | None = None) -> str:
        """Lưu workbook, trả về path đã lưu (save-as nếu truyền path mới)."""
