# MCP Server Plan: Real-Time Excel Formula Tracing & Label Identification

**Audience:** AI Agent

**Version:** 1.0

## 1. Objective

Enable an AI to interact with an active Microsoft Excel instance in real-time to:

*   Trace formula precedents and dependents.
*   Identify and map human-readable labels to cell addresses.
*   Understand the structure and data flow within an Excel sheet.

## 2. Core Technology Stack

*   **Language:** Python
*   **Excel Interaction:** `pywin32` library for COM-based automation of Microsoft Excel.
*   **Server Framework:** A lightweight Python framework for exposing MCP endpoints (FastMCP - https://gofastmcp.com/getting-started/welcome).

## 3. Key MCP Server Functions/Endpoints

*   **`initialize_excel_link()`**
    *   **Purpose:** Establish a connection to a running Excel instance or open a specified workbook.
    *   **Action:** Uses `pywin32` to get or create an Excel Application object.
    *   **Returns:** Status (success/failure), Workbook name, Active Sheet name.

*   **`get_formula(sheet_name, cell_address)`**
    *   **Purpose:** Retrieve the formula from a specified cell.
    *   **Inputs:** `sheet_name` (string), `cell_address` (string, e.g., "A1").
    *   **Action:** Accesses `worksheet.Range(cell_address).Formula`.
    *   **Returns:** Formula string (or value if no formula).

*   **`trace_precedents(sheet_name, cell_address)`**
    *   **Purpose:** Identify all cells that directly or indirectly feed into the formula of a given cell.
    *   **Inputs:** `sheet_name` (string), `cell_address` (string).
    *   **Action:** Uses `Range.Precedents` property. May require recursive calls for multi-level precedents.
    *   **Returns:** List of precedent cell addresses (e.g., `["Sheet1!B2", "Sheet1!C3"]`).

*   **`trace_dependents(sheet_name, cell_address)`**
    *   **Purpose:** Identify all cells whose formulas directly or indirectly use the value of a given cell.
    *   **Inputs:** `sheet_name` (string), `cell_address` (string).
    *   **Action:** Uses `Range.Dependents` property. May require recursive calls for multi-level dependents.
    *   **Returns:** List of dependent cell addresses.

*   **`find_cell_labels(sheet_name, cell_address, search_radius=1)`**
    *   **Purpose:** Attempt to identify a human-readable label for a given data cell.
    *   **Inputs:**
        *   `sheet_name` (string).
        *   `cell_address` (string).
        *   `search_radius` (integer, optional): Number of adjacent cells (left, right, top, bottom) to check for text values that might serve as labels.
    *   **Action:**
        *   Checks cells immediately adjacent (e.g., to the left or above) the target cell for non-numeric, non-formula text content.
        *   Consider checking Excel Named Ranges that include the `cell_address`.
    *   **Returns:** A list of potential labels found (e.g., `["Sales Revenue", "FY2024 Actual"]`) or an empty list.

*   **`build_label_address_map(sheet_name, scan_range=None)`**
    *   **Purpose:** Scan a specified range or an entire sheet to create a heuristic mapping of potential labels to their corresponding data cell addresses.
    *   **Inputs:**
        *   `sheet_name` (string).
        *   `scan_range` (string, optional, e.g., "A1:Z100"): The area to scan. If `None`, attempts to scan the used range of the sheet.
    *   **Action:**
        *   Iterate through cells in `scan_range`.
        *   For each cell containing text, check adjacent cells (typically to its right or below) for numerical data or formulas.
        *   If a pattern (text label next to data/formula) is found, add it to the map.
        *   Incorporate Excel Named Ranges into this map.
    *   **Returns:** A dictionary where keys are identified labels (string) and values are cell addresses (string), e.g., `{"Net Income": "B20", "Total Assets": "F15"}`.

## 4. Real-Time Considerations (Event Handling)

*   `pywin32` allows for handling Excel events (e.g., `Application.SheetChange`, `Workbook.SheetCalculate`).
*   The MCP server could subscribe to these events.
*   Upon a relevant event (e.g., a cell value or formula change), the server can:
    *   Notify the AI client.
    *   Re-run specific tracing or label-finding functions for affected areas.
    *   Update its internal cache of label-address maps.
*   **Challenge:** Event handling can be complex and resource-intensive. A polling mechanism might be a simpler initial approach if full real-time eventing is difficult.

## 5. Data Storage (Internal to MCP Server)

*   **Local Database:** Utilize DuckDB for storing and querying the `label_address_map` and potentially other structured data derived from Excel.
    *   **Mode:** Can be run in-memory for speed or persisted to a file for data retention across server restarts (though frequent updates from live Excel data are expected).
    *   **Table Structure (Example):** A table like `cell_labels (label TEXT, sheet_name TEXT, cell_address TEXT, last_updated TIMESTAMP)`.
*   **Querying:** Use SQL for efficient lookups and potential analysis of the stored data.
*   **Cache Invalidation:** Updates from Excel (via events or polling) will trigger updates or replacements of records in the DuckDB table(s).

## 6. Assumptions for the Consuming AI

*   The AI understands basic Excel grid concepts (sheets, rows, columns, cell addresses).
*   The AI can make HTTP/RPC calls to the MCP server endpoints.
*   The AI can interpret the structured data (JSON) returned by the server.