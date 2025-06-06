from typing import Optional, List, Set

try:
    import win32com.client as win32
except ImportError:  # Not on Windows or pywin32 not installed
    win32 = None

from fastmcp.server import FastMCP

# FastMCP server instance
server = FastMCP(name="excel-mcp")

excel_app = None

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


def _is_text_label(cell) -> bool:
    """Return True if a cell contains a potential text label."""
    try:
        value = cell.Value
    except Exception:
        return False

    if value is None:
        return False

    # Exclude formulas or numeric values
    if getattr(cell, "HasFormula", False):
        return False

    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return False
        try:
            float(stripped.replace(",", ""))
            return False
        except ValueError:
            return True

    return False


def _adjacent_label_search(ws, row: int, col: int, radius: int = 1) -> List[str]:
    """Search nearby cells for text labels."""
    labels: List[str] = []
    directions = [
        (0, -1),  # left
        (0, 1),   # right
        (-1, 0),  # up
        (1, 0),   # down
    ]

    max_row = ws.Rows.Count
    max_col = ws.Columns.Count

    for dr, dc in directions:
        for i in range(1, radius + 1):
            r = row + dr * i
            c = col + dc * i
            if r < 1 or c < 1 or r > max_row or c > max_col:
                break
            cell = ws.Cells(r, c)
            if _is_text_label(cell):
                labels.append(str(cell.Value).strip())
                break
    return labels


def _named_range_labels(ws, cell) -> List[str]:
    """Return labels from named ranges referring to the cell."""
    labels: List[str] = []
    addr = cell.Address
    wb = ws.Parent

    # Workbook-level names
    for name in wb.Names:
        try:
            rng = name.RefersToRange
        except Exception:
            continue
        try:
            if rng.Worksheet.Name == ws.Name and rng.Address == addr:
                labels.append(name.Name)
        except Exception:
            continue

    # Sheet-level names
    try:
        for name in ws.Names:
            try:
                rng = name.RefersToRange
            except Exception:
                continue
            if rng.Address == addr:
                labels.append(name.Name)
    except Exception:
        pass

    return labels


@server.tool
def find_cell_labels(sheet_name: Optional[str], cell_address: str, search_radius: int = 1):
    """Attempt to identify text labels describing the given cell."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        target = ws.Range(cell_address)

        labels: List[str] = []
        labels.extend(_adjacent_label_search(ws, target.Row, target.Column, search_radius))
        labels.extend(_named_range_labels(ws, target))

        # Deduplicate while preserving order
        seen = set()
        deduped = []
        for lbl in labels:
            if lbl not in seen:
                seen.add(lbl)
                deduped.append(lbl)

        return {
            "status": "success",
            "sheet": ws.Name,
            "address": cell_address,
            "labels": deduped,
        }
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


if __name__ == "__main__":
    # Run server using HTTP transport by default
    server.run(transport="streamable-http")
