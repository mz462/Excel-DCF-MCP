from typing import Optional, List, Set, Dict

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


@server.tool
def build_label_address_map(sheet_name: Optional[str], scan_range: Optional[str] = None):
    """Return a mapping of text labels to nearby data cell addresses."""
    if win32 is None:
        return {"status": "failure", "reason": "pywin32 not available"}

    if excel_app is None:
        return {"status": "failure", "reason": "excel link not initialized"}

    try:
        wb = excel_app.ActiveWorkbook
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet

        scan_rng = ws.Range(scan_range) if scan_range else ws.UsedRange

        label_map: Dict[str, str] = {}

        for cell in scan_rng:
            if isinstance(cell.Value, str) and cell.Value.strip():
                label = cell.Value.strip()
                target = None
                try:
                    right = cell.Offset(0, 1)
                    if right.Value not in (None, "") and not isinstance(right.Value, str):
                        target = right
                except Exception:
                    pass
                if target is None:
                    try:
                        below = cell.Offset(1, 0)
                        if below.Value not in (None, "") and not isinstance(below.Value, str):
                            target = below
                    except Exception:
                        pass

                if target is not None:
                    addr = f"{target.Worksheet.Name}!{target.Address(False, False)}"
                    label_map[label] = addr

        # Include named ranges
        for name in wb.Names:
            try:
                refers = name.RefersToRange
                if refers.Worksheet.Name == ws.Name:
                    addr = f"{refers.Worksheet.Name}!{refers.Address(False, False)}"
                    label_map[name.Name] = addr
            except Exception:
                continue

        return {"status": "success", "sheet": ws.Name, "label_map": label_map}
    except Exception as e:
        return {"status": "failure", "reason": str(e)}


if __name__ == "__main__":
    # Run server using HTTP transport by default
    server.run(transport="streamable-http")
