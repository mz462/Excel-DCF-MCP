# Excel MCP Server

This project exposes **FastMCP** tools to interact with a running Microsoft Excel instance. The server is intended to be called by an AI agent.

Currently implemented:
- `initialize_excel_link` tool to connect to Excel (requires Windows with pywin32).
- `get_formula` tool to read formulas or values from cells.
- `trace_precedents` tool to list all precedent cells for a target cell.
- `trace_dependents` tool to list all cells that depend on a target cell.

Run the server:
```
python -m excel_mcp.server
```
