from typing import Optional, List, Set
import threading
import time

from . import db

try:
    import pythoncom  # type: ignore
except ImportError:  # Not on Windows or pywin32 not installed
    pythoncom = None

try:
    import win32com.client as win32
except ImportError:  # Not on Windows or pywin32 not installed
    win32 = None

from fastmcp.server import FastMCP

# FastMCP server instance
server = FastMCP(name="excel-mcp")

excel_app = None
_excel_event_handler = None
_event_thread = None
_event_stop = threading.Event()


class _ExcelEventSink:
    """Simple event sink for Excel Application events."""

    def __init__(self):
        self.events: List[dict] = []

    def OnSheetChange(self, sh, target):  # pylint: disable=invalid-name
        addr = f"{sh.Name}!{target.Address(False, False)}"
        self.events.append({
            "event": "SheetChange",
            "sheet": sh.Name,
            "address": addr,
        })

    def OnSheetCalculate(self, sh):  # pylint: disable=invalid-name
        self.events.append({
            "event": "SheetCalculate",
            "sheet": sh.Name,
        })


def _event_loop():
    if pythoncom is None:
        return
    while not _event_stop.is_set():
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)

@server.tool
def initialize_database(path: str = "excel_mcp.db"):
    """Initialize persistent DuckDB storage."""
    try:
        db.init_db(path)
        return {"status": "success", "path": path}
    except Exception as e:  # pragma: no cover - simple wrapper
        return {"status": "failure", "reason": str(e)}

@server.tool
def initialize_excel_link(workbook: Optional[str] = None):
    """Establish a connection to a running Excel instance or open a workbook."""
    global excel_app
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    try:
        excel_app = win32.GetActiveObject("Excel.Application")
    except Exception:
        excel_app = win32.Dispatch("Excel.Application")

    if workbook:
        excel_app.Workbooks.Open(workbook)

    excel_app.Visible = True
    wb = excel_app.ActiveWorkbook
    ws = wb.ActiveSheet
    return {
        "status": "success",
        "workbook": wb.Name,
        "sheet": ws.Name,
    }


@server.tool
def get_formula(sheet_name: Optional[str], cell_address: str):
    """Return the formula from a cell or the value if no formula exists."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        cell = ws.Range(cell_address)
        formula = cell.Formula
        if formula == "":
            return {
                "status": "success",
                "sheet": ws.Name,
                "address": cell_address,
                "value": cell.Value,
            }
        return {
            "status": "success",
            "sheet": ws.Name,
            "address": cell_address,
            "formula": formula,
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


@server.tool
def trace_precedents(sheet_name: Optional[str], cell_address: str):
    """Return all precedent cell addresses for a given cell."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        start_cell = ws.Range(cell_address)

        addresses: Set[str] = set()

        def _collect_precedents(rng):
            try:
                precs = rng.Precedents
            except Exception:
                return
            try:
                for cell in precs:
                    addr = f"{cell.Worksheet.Name}!{cell.Address(False, False)}"
                    if addr not in addresses:
                        addresses.add(addr)
                        _collect_precedents(cell)
            except TypeError:
                # If precs is a single cell range, iteration may fail
                cell = precs
                addr = f"{cell.Worksheet.Name}!{cell.Address(False, False)}"
                if addr not in addresses:
                    addresses.add(addr)
                    _collect_precedents(cell)

        _collect_precedents(start_cell)

        return {
            "status": "success",
            "sheet": ws.Name,
            "address": cell_address,
            "precedents": sorted(addresses),
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


@server.tool
def trace_dependents(sheet_name: Optional[str], cell_address: str):
    """Return all dependent cell addresses for a given cell."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        start_cell = ws.Range(cell_address)

        addresses: Set[str] = set()

        def _collect_dependents(rng):
            try:
                deps = rng.Dependents
            except Exception:
                return
            try:
                for cell in deps:
                    addr = f"{cell.Worksheet.Name}!{cell.Address(False, False)}"
                    if addr not in addresses:
                        addresses.add(addr)
                        _collect_dependents(cell)
            except TypeError:
                # If deps is a single cell range, iteration may fail
                cell = deps
                addr = f"{cell.Worksheet.Name}!{cell.Address(False, False)}"
                if addr not in addresses:
                    addresses.add(addr)
                    _collect_dependents(cell)

        _collect_dependents(start_cell)

        return {
            "status": "success",
            "sheet": ws.Name,
            "address": cell_address,
            "dependents": sorted(addresses),
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


def _adjacent_text_labels(cell, radius: int) -> List[str]:
    """Return text values in cells near the target cell."""
    labels: List[str] = []
    for row_offset in range(-radius, radius + 1):
        for col_offset in range(-radius, radius + 1):
            if row_offset == 0 and col_offset == 0:
                continue
            try:
                adj = cell.Offset(row_offset, col_offset)
                if adj.Formula == "" and isinstance(adj.Value, str) and adj.Value.strip() != "":
                    labels.append(str(adj.Value).strip())
            except Exception:
                continue
    return labels


@server.tool
def find_cell_labels(sheet_name: Optional[str], cell_address: str, search_radius: int = 1):
    """Attempt to identify human-readable labels for a given cell."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        cell = ws.Range(cell_address)

        labels: List[str] = []

        # Adjacent text values
        labels.extend(_adjacent_text_labels(cell, search_radius))

        # Named ranges including the cell
        try:
            for n in wb.Names:
                try:
                    rng = n.RefersToRange
                except Exception:
                    continue
                try:
                    if cell.Address in rng.Address:
                        labels.append(n.Name)
                except Exception:
                    continue
        except Exception:
            pass

        return {
            "status": "success",
            "sheet": ws.Name,
            "address": cell_address,
            "labels": sorted(set(labels)),
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


@server.tool
def build_label_address_map(sheet_name: Optional[str], scan_range: Optional[str] = None):
    """Return a heuristic mapping of labels to cell addresses for a worksheet."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet

        rng = ws.Range(scan_range) if scan_range else ws.UsedRange

        label_map = {}

        first_row = rng.Row
        first_col = rng.Column
        rows = rng.Rows.Count
        cols = rng.Columns.Count

        for r in range(rows):
            for c in range(cols):
                cell = ws.Cells(first_row + r, first_col + c)
                if cell.Formula == "" and isinstance(cell.Value, str) and cell.Value.strip() != "":
                    label = str(cell.Value).strip()
                    right = cell.Offset(0, 1)
                    below = cell.Offset(1, 0)
                    target = None
                    try:
                        if right.Formula != "" or (right.Formula == "" and isinstance(right.Value, (int, float))):
                            target = right
                    except Exception:
                        pass
                    if target is None:
                        try:
                            if below.Formula != "" or (below.Formula == "" and isinstance(below.Value, (int, float))):
                                target = below
                        except Exception:
                            pass
                    if target is not None:
                        addr = f"{target.Worksheet.Name}!{target.Address(False, False)}"
                        if label not in label_map:
                            label_map[label] = addr

        # Include named ranges
        try:
            for n in wb.Names:
                try:
                    cell = n.RefersToRange.Cells(1, 1)
                    addr = f"{cell.Worksheet.Name}!{cell.Address(False, False)}"
                    if n.Name not in label_map:
                        label_map[n.Name] = addr
                except Exception:
                    continue
        except Exception:
            pass

        db.store_label_map(ws.Name, label_map)
        return {
            "status": "success",
            "sheet": ws.Name,
            "label_map": label_map,
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


@server.tool
def query_label(label: str):
    """Query stored label mappings from the database."""
    try:
        rows = db.query_label(label)
        results = [{"sheet": r[0], "address": r[1]} for r in rows]
        return {"status": "success", "results": results}
    except Exception as e:  # pragma: no cover - simple wrapper
        return {"status": "failure", "reason": str(e)}


@server.tool
def start_excel_event_monitor():
    """Begin monitoring Excel events to record changes."""
    global _excel_event_handler, _event_thread
    if win32 is None or pythoncom is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    if _event_thread and _event_thread.is_alive():
        return {"status": "running"}

    _excel_event_handler = win32.WithEvents(excel_app, _ExcelEventSink)
    _event_stop.clear()
    _event_thread = threading.Thread(target=_event_loop, daemon=True)
    _event_thread.start()
    return {"status": "started"}


@server.tool
def stop_excel_event_monitor():
    """Stop monitoring Excel events."""
    global _excel_event_handler, _event_thread
    if _event_thread is None:
        return {"status": "not_running"}

    _event_stop.set()
    _event_thread.join()
    _event_thread = None
    _excel_event_handler = None
    return {"status": "stopped"}


@server.tool
def fetch_excel_events():
    """Retrieve and clear recorded Excel events."""
    if _excel_event_handler is None:
        return {"status": "failure", "reason": "event monitor not running"}

    events = list(_excel_event_handler.events)
    _excel_event_handler.events.clear()
    return {"status": "success", "events": events}


if __name__ == "__main__":
    # Run server using HTTP transport by default
    server.run(transport="streamable-http")
