# Excel MCP Server

This project exposes **FastMCP** tools to interact with a running Microsoft Excel instance. The server is intended to be called by an AI agent.

Currently implemented:
- `initialize_excel_link` tool to connect to Excel (requires Windows with pywin32).
- `get_formula` tool to read formulas or values from cells.
- `trace_precedents` tool to list all precedent cells for a target cell.
- `trace_dependents` tool to list all cells that depend on a target cell.
- `find_cell_labels` tool to guess human-readable labels for a cell.
- `build_label_address_map` tool to map labels to data cell addresses.
- Excel event monitoring tools to capture cell changes.
- DuckDB persistence for label mappings (`initialize_database`, `query_label`).

Run the server:
```
python -m excel_mcp.server
```

## Running Tests

Unit tests are located in the `tests` directory and can be executed with:

```
python -m unittest discover -s tests
```

## Example

A minimal example script is available in `examples/basic_usage.py` which starts
the MCP server using the default HTTP transport.
