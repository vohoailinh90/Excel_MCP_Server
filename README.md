# Excel MCP Server — v1

An MCP server that exposes Excel operations (open/close files, read/write
cells, create sheets, apply formulas, and **run existing VBA macros by
name**) so an AI agent (Claude Desktop, Claude Code, or any MCP client)
can call them.

## Already tested in this sandbox

```
python demo/demo_test.py
```

Result: all 12 tools are registered correctly by the server, and the
full flow `open → read/write cell → apply formula → create sheet
(copy) → run macro → save → read back from disk to verify` runs
correctly end-to-end using `MockOpenpyxlBackend`.

**Important limitation of this demo**: it runs on Linux (no Excel
installed), so the "macro" in the demo is a Python function registered
manually to mimic the call-by-name convention — **it is not real VBA**.
To run your actual `.xlsm` macros, this needs to run on a Windows
machine with Excel installed (see below).

## Architecture

```
                      ┌─────────────────────┐
  Claude / AI agent   │   MCP Client         │
  (Desktop, Code,...) │  (stdio transport)   │
                      └──────────┬──────────┘
                                 │ JSON-RPC over stdio
                      ┌──────────▼──────────┐
                      │   server.py          │  ← 12 tools: excel_open_workbook,
                      │   (mcp.MCPServer)     │    excel_read_cell, excel_run_macro, ...
                      └──────────┬──────────┘
                                 │ calls through a shared interface
                      ┌──────────▼──────────┐
                      │  backends/base.py     │  (ABC, OS-independent)
                      └──────────┬──────────┘
                     ┌───────────┴────────────┐
        ┌────────────▼───────────┐  ┌─────────▼────────────┐
        │ windows_com.py          │  │ mock_openpyxl.py       │
        │ (xlwings + pywin32)     │  │ (openpyxl, any OS)     │
        │ - drives real Excel.exe │  │ - test/dev only, does  │
        │ - runs real VBA macros  │  │   NOT run real VBA     │
        │   in .xlsm files        │  └────────────────────────┘
        │ - dedicated COM thread  │
        │   to avoid STA errors   │
        └─────────────────────────┘
```

`server.py` automatically picks the backend: `platform.system() ==
"Windows"` → `WindowsComBackend`, otherwise → `MockOpenpyxlBackend`
(with a warning printed to stderr).

### Why xlwings instead of raw win32com?

You already have experience using `pywin32` directly (the Outlook RSVP
Tool). For Excel, `xlwings` is a better fit because:
- `wb.macro("MacroName")(*args)` calls a VBA `Sub`/`Function` directly
  by name, whether it lives in a regular Module or in `ThisWorkbook`.
- The cell/range API (`sheet.range("A1").value`) is much cleaner than
  raw COM's `worksheet.Cells(1,1).Value`.
- It still uses pywin32 under the hood, so you don't lose the ability
  to "drop down" to raw COM (`wb.api`, `sheet.api`) for advanced
  operations (pivot tables, charts, the VBA editor...) that xlwings
  doesn't wrap yet.

### Why a dedicated COM thread?

Excel COM is single-threaded apartment (STA). If multiple tool calls
arrive close together (an agent may call tools asynchronously), calling
COM directly from different threads/coroutines will crash or produce
"Application is busy" errors. The fix: a single background thread holds
`CoInitialize()`, and every command is pushed into a queue and executed
sequentially on that same thread — safe, simple, and fast enough for
the expected scale ("a handful of agents calling tools", not hundreds
of requests per second).

## Installation

### On a dev/test machine (no Windows required)

```bash
pip install -r requirements.txt
python demo/demo_test.py
```

### On a production Windows machine (with Excel + real macros)

```powershell
pip install -r requirements.txt
python server.py
```

The server will automatically detect Windows and use
`WindowsComBackend`.

**Safety recommendation**: restrict which macros can be run via an
environment variable, so an AI agent can't call arbitrary macros on a
real file:

```powershell
$env:EXCEL_MCP_ALLOWED_MACROS = "GenerateSummary,RefreshPivot,ExportReport"
python server.py
```

## Configuring Claude Desktop

Add this to Claude Desktop's MCP config file
(`%APPDATA%\Claude\claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "excel-automation": {
      "command": "python",
      "args": ["C:\\path\\to\\excel-mcp-server\\server.py"],
      "env": {
        "EXCEL_MCP_ALLOWED_MACROS": "GenerateSummary,RefreshPivot"
      }
    }
  }
}
```

For Claude Code, add it the same way to the project's MCP server
config (via `.claude` or the `claude mcp add` command).

## Tool list (v1)

| Tool | What it does |
|---|---|
| `excel_open_workbook(path, visible)` | Opens a file, returns `workbook_id` |
| `excel_close_workbook(workbook_id, save)` | Closes a file |
| `excel_list_open_workbooks()` | Lists currently open files |
| `excel_list_sheets(workbook_id)` | Lists sheet names |
| `excel_read_cell(workbook_id, sheet, cell)` | Reads one cell |
| `excel_read_range(workbook_id, sheet, cell_range)` | Reads a range |
| `excel_write_cell(workbook_id, sheet, cell, value)` | Writes one cell |
| `excel_write_range(workbook_id, sheet, start_cell, values)` | Writes a 2D array |
| `excel_apply_formula(workbook_id, sheet, cell, formula)` | Applies a formula |
| `excel_create_sheet(workbook_id, name, copy_from, position)` | Creates/copies a sheet |
| `excel_run_macro(workbook_id, macro_name, args)` | Runs a VBA macro by name |
| `excel_save_workbook(workbook_id, path)` | Saves / save-as |

## Next steps (not part of this v1)

1. **Test with your actual `.xlsm` macros** on a Windows machine — need
   the specific Sub/Function names and argument counts to verify
   `excel_run_macro` works as expected.
2. **Handling multiple workbooks concurrently** more carefully —
   currently a single shared `xlwings.App` is used; if full isolation
   is needed (e.g. 2 agents working on 2 independent files without
   affecting each other's global Excel settings), separate `App`
   instances could be used instead — at the cost of more RAM/CPU.
3. **Idle timeout**: automatically close workbooks and shut down the
   `App` if no command arrives within N minutes, to avoid a hidden
   Excel process silently eating RAM if an agent forgets to close it.
4. **Logging/audit**: log every write/macro command (who called it,
   when, which macro, which arguments) — important since macros are
   arbitrary code execution.
5. **Retry/file-lock handling**: handle the case where you already have
   the file open manually in Excel while the agent also tries to open
   it — currently this raises a COM error that should be caught and
   turned into a clear message.
6. Consider `streamable-http` transport instead of `stdio` if you want
   multiple different clients/agents to connect to one long-running
   server process, instead of each Claude Desktop instance spawning its
   own server process.
