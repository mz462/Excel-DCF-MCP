from typing import Optional

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


if __name__ == "__main__":
    # Run server using HTTP transport by default
    server.run(transport="streamable-http")
