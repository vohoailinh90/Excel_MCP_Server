"""
Excel MCP Server - v1 (bản đầu tiên, thực dụng)

Chạy: python server.py
Cấu hình để Claude Desktop / Claude Code gọi tới: xem README.md

Chọn backend:
- Windows (có Excel cài sẵn) -> WindowsComBackend (xlwings), chạy được
  macro VBA thật trong file .xlsm.
- Mọi OS khác (demo/dev) -> MockOpenpyxlBackend, không chạy được macro
  VBA thật, chỉ demo luồng gọi tool.

An toàn: macro là code tuỳ ý (arbitrary code execution). KHÔNG nên để
AI agent gọi bất kỳ macro nào theo tên nó tự nghĩ ra trên workbook sản
xuất thật. Set biến môi trường EXCEL_MCP_ALLOWED_MACROS (phân cách bởi
dấu phẩy) để giới hạn danh sách macro được phép chạy.
"""

from __future__ import annotations

import os
import platform
import sys
from typing import Any

from mcp.server.mcpserver import MCPServer

from backends.base import ExcelBackend, ExcelBackendError

# ---- chọn backend theo OS ----
_allowed_env = os.environ.get("EXCEL_MCP_ALLOWED_MACROS", "")
_allowed_macros = {m.strip() for m in _allowed_env.split(",") if m.strip()} or None

backend: ExcelBackend
if platform.system() == "Windows":
    from backends.windows_com import WindowsComBackend

    backend = WindowsComBackend(allowed_macros=_allowed_macros)
else:
    from backends.mock_openpyxl import MockOpenpyxlBackend

    backend = MockOpenpyxlBackend()
    print(
        "[excel-mcp-server] Không chạy trên Windows -> dùng MockOpenpyxlBackend. "
        "Macro VBA thật sẽ KHÔNG chạy được. Xem README.md.",
        file=sys.stderr,
    )

mcp = MCPServer(
    name="excel-automation",
    version="0.1.0",
    instructions=(
        "Công cụ điều khiển Excel: mở/đóng file, đọc/ghi cell, tạo sheet, "
        "áp công thức, và chạy macro VBA có sẵn theo tên. Luôn gọi "
        "excel_open_workbook trước, dùng workbook_id trả về cho các lệnh "
        "tiếp theo, và nhớ gọi excel_close_workbook hoặc excel_save_workbook "
        "khi xong để không mất dữ liệu."
    ),
)


@mcp.tool()
def excel_open_workbook(path: str, visible: bool = False) -> dict[str, Any]:
    """Mở 1 file Excel (.xlsx/.xlsm) và trả về workbook_id để dùng cho các lệnh sau."""
    try:
        wb_id = backend.open_workbook(path, visible=visible)
        return {"workbook_id": wb_id, "path": path}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_close_workbook(workbook_id: str, save: bool = True) -> dict[str, Any]:
    """Đóng workbook đang mở. save=True sẽ lưu trước khi đóng."""
    try:
        backend.close_workbook(workbook_id, save=save)
        return {"status": "closed", "saved": save}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_list_open_workbooks() -> list[dict[str, Any]]:
    """Liệt kê các workbook đang mở qua server này."""
    return backend.list_open_workbooks()


@mcp.tool()
def excel_list_sheets(workbook_id: str) -> dict[str, Any]:
    """Liệt kê tên các sheet trong workbook."""
    try:
        return {"sheets": backend.list_sheets(workbook_id)}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_read_cell(workbook_id: str, sheet: str, cell: str) -> dict[str, Any]:
    """Đọc giá trị 1 ô, ví dụ cell='B5'."""
    try:
        return {"value": backend.read_cell(workbook_id, sheet, cell)}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_read_range(workbook_id: str, sheet: str, cell_range: str) -> dict[str, Any]:
    """Đọc 1 vùng ô, ví dụ cell_range='A1:C10'. Trả về mảng 2 chiều."""
    try:
        return {"values": backend.read_range(workbook_id, sheet, cell_range)}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_write_cell(workbook_id: str, sheet: str, cell: str, value: Any) -> dict[str, Any]:
    """Ghi 1 giá trị vào 1 ô."""
    try:
        backend.write_cell(workbook_id, sheet, cell, value)
        return {"status": "ok"}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_write_range(
    workbook_id: str, sheet: str, start_cell: str, values: list[list[Any]]
) -> dict[str, Any]:
    """Ghi 1 mảng 2 chiều bắt đầu từ start_cell (ví dụ 'A1')."""
    try:
        backend.write_range(workbook_id, sheet, start_cell, values)
        return {"status": "ok"}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_apply_formula(workbook_id: str, sheet: str, cell: str, formula: str) -> dict[str, Any]:
    """Gán công thức vào 1 ô, ví dụ formula='=SUM(A1:A10)'."""
    try:
        backend.apply_formula(workbook_id, sheet, cell, formula)
        return {"status": "ok"}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_create_sheet(
    workbook_id: str,
    name: str,
    copy_from: str | None = None,
    position: int | None = None,
) -> dict[str, Any]:
    """Tạo sheet mới. Nếu truyền copy_from, sheet mới sẽ là bản sao của sheet đó."""
    try:
        backend.create_sheet(workbook_id, name, copy_from=copy_from, position=position)
        return {"status": "ok", "name": name}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_run_macro(
    workbook_id: str, macro_name: str, args: list[Any] | None = None
) -> dict[str, Any]:
    """Chạy 1 macro VBA có sẵn trong workbook theo tên, truyền args nếu macro cần tham số."""
    try:
        result = backend.run_macro(workbook_id, macro_name, args=args)
        return {"status": "ok", "result": result}
    except ExcelBackendError as e:
        return {"error": str(e)}


@mcp.tool()
def excel_save_workbook(workbook_id: str, path: str | None = None) -> dict[str, Any]:
    """Lưu workbook. Nếu truyền path, sẽ save-as vào path đó."""
    try:
        saved_path = backend.save_workbook(workbook_id, path=path)
        return {"status": "ok", "path": saved_path}
    except ExcelBackendError as e:
        return {"error": str(e)}


if __name__ == "__main__":
    mcp.run(transport="stdio")
