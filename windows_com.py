"""
WindowsComBackend: điều khiển Excel thật qua COM (xlwings) để chạy macro
VBA có sẵn, đọc/ghi cell, tạo sheet...

QUAN TRỌNG VỀ THREADING:
Excel COM là single-threaded apartment (STA). Nếu MCP server nhận nhiều
tool-call đồng thời và gọi COM từ nhiều thread/coroutine khác nhau, Excel
sẽ crash hoặc trả lỗi khó hiểu. Giải pháp: một thread nền DUY NHẤT,
CoInitialize() một lần trên thread đó, mọi lệnh Excel được đẩy vào hàng
đợi (queue) và thực thi tuần tự trên đúng thread ấy. server.py (async)
gọi qua `submit()` và await kết quả bằng run_in_executor kiểu 1-worker.

Cài đặt (trên máy Windows của anh):
    pip install xlwings pywin32
"""

from __future__ import annotations

import queue
import threading
import uuid
from typing import Any, Callable

from .base import ExcelBackend, ExcelBackendError


class _ComWorker(threading.Thread):
    """Thread nền duy nhất, sở hữu COM apartment, chạy job tuần tự."""

    def __init__(self) -> None:
        super().__init__(daemon=True, name="excel-com-worker")
        self._jobs: "queue.Queue[tuple[Callable, tuple, dict, queue.Queue]]" = queue.Queue()
        self._ready = threading.Event()
        self.start()
        self._ready.wait()

    def run(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        self._ready.set()
        try:
            while True:
                fn, args, kwargs, result_q = self._jobs.get()
                if fn is None:  # sentinel để dừng thread
                    break
                try:
                    result_q.put(("ok", fn(*args, **kwargs)))
                except Exception as e:  # noqa: BLE001
                    result_q.put(("error", e))
        finally:
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable, *args, **kwargs) -> Any:
        result_q: "queue.Queue[tuple[str, Any]]" = queue.Queue()
        self._jobs.put((fn, args, kwargs, result_q))
        status, value = result_q.get()
        if status == "error":
            raise value
        return value

    def stop(self) -> None:
        self._jobs.put((None, (), {}, queue.Queue()))


class WindowsComBackend(ExcelBackend):
    def __init__(self, allowed_macros: set[str] | None = None) -> None:
        """
        allowed_macros: nếu truyền vào, run_macro() chỉ cho phép chạy các
        tên macro nằm trong danh sách này. Rất nên dùng trong production -
        không nên để AI agent chạy macro tuỳ ý theo tên do nó tự nghĩ ra.
        """
        self._worker = _ComWorker()
        self._workbooks: dict[str, Any] = {}  # workbook_id -> xlwings.Book
        self._paths: dict[str, str] = {}
        self._allowed_macros = allowed_macros
        self._app = None

    # ---- helpers chạy trên COM thread ----
    def _ensure_app(self):
        import xlwings as xw

        if self._app is None:
            self._app = xw.App(visible=False, add_book=False)
        return self._app

    def _get_wb(self, workbook_id: str):
        wb = self._workbooks.get(workbook_id)
        if wb is None:
            raise ExcelBackendError(f"workbook_id '{workbook_id}' không tồn tại hoặc đã đóng")
        return wb

    # ---- API public (mỗi hàm submit 1 job vào COM thread) ----

    def open_workbook(self, path: str, visible: bool = False) -> str:
        def _do():
            app = self._ensure_app()
            app.visible = visible
            wb = app.books.open(path)
            wb_id = str(uuid.uuid4())[:8]
            self._workbooks[wb_id] = wb
            self._paths[wb_id] = path
            return wb_id

        return self._worker.submit(_do)

    def close_workbook(self, workbook_id: str, save: bool = True) -> None:
        def _do():
            wb = self._get_wb(workbook_id)
            if save:
                wb.save()
            wb.close()
            del self._workbooks[workbook_id]
            del self._paths[workbook_id]

        self._worker.submit(_do)

    def list_open_workbooks(self) -> list[dict[str, Any]]:
        return [{"workbook_id": k, "path": v} for k, v in self._paths.items()]

    def list_sheets(self, workbook_id: str) -> list[str]:
        def _do():
            wb = self._get_wb(workbook_id)
            return [s.name for s in wb.sheets]

        return self._worker.submit(_do)

    def read_cell(self, workbook_id: str, sheet: str, cell: str) -> Any:
        def _do():
            wb = self._get_wb(workbook_id)
            return wb.sheets[sheet].range(cell).value

        return self._worker.submit(_do)

    def read_range(self, workbook_id: str, sheet: str, cell_range: str) -> list[list[Any]]:
        def _do():
            wb = self._get_wb(workbook_id)
            val = wb.sheets[sheet].range(cell_range).value
            # xlwings trả về scalar nếu range 1 ô, list phẳng nếu 1 hàng/cột
            if not isinstance(val, list):
                return [[val]]
            if val and not isinstance(val[0], list):
                return [val]
            return val

        return self._worker.submit(_do)

    def write_cell(self, workbook_id: str, sheet: str, cell: str, value: Any) -> None:
        def _do():
            wb = self._get_wb(workbook_id)
            wb.sheets[sheet].range(cell).value = value

        self._worker.submit(_do)

    def write_range(
        self, workbook_id: str, sheet: str, start_cell: str, values: list[list[Any]]
    ) -> None:
        def _do():
            wb = self._get_wb(workbook_id)
            wb.sheets[sheet].range(start_cell).value = values

        self._worker.submit(_do)

    def apply_formula(self, workbook_id: str, sheet: str, cell: str, formula: str) -> None:
        def _do():
            wb = self._get_wb(workbook_id)
            wb.sheets[sheet].range(cell).formula = formula

        self._worker.submit(_do)

    def create_sheet(
        self,
        workbook_id: str,
        name: str,
        copy_from: str | None = None,
        position: int | None = None,
    ) -> None:
        def _do():
            wb = self._get_wb(workbook_id)
            if copy_from:
                src = wb.sheets[copy_from]
                src.api.Copy(After=wb.sheets[-1].api)
                new_sheet = wb.sheets[-1]
                new_sheet.name = name
            else:
                after = wb.sheets[-1] if position is None else wb.sheets[position - 1]
                wb.sheets.add(name=name, after=after)

        self._worker.submit(_do)

    def run_macro(self, workbook_id: str, macro_name: str, args: list[Any] | None = None) -> Any:
        if self._allowed_macros is not None and macro_name not in self._allowed_macros:
            raise ExcelBackendError(
                f"Macro '{macro_name}' không nằm trong allow-list. "
                f"Cho phép: {sorted(self._allowed_macros)}"
            )

        def _do():
            wb = self._get_wb(workbook_id)
            macro = wb.macro(macro_name)
            return macro(*(args or []))

        return self._worker.submit(_do)

    def save_workbook(self, workbook_id: str, path: str | None = None) -> str:
        def _do():
            wb = self._get_wb(workbook_id)
            if path:
                wb.save(path)
                self._paths[workbook_id] = path
            else:
                wb.save()
            return self._paths[workbook_id]

        return self._worker.submit(_do)
