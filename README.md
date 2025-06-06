# Excel MCP Server

This project exposes **FastMCP** tools to interact with a running Microsoft Excel instance. The server is intended to be called by an AI agent.

Currently implemented:
- `initialize_excel_link` tool to connect to Excel (requires Windows with pywin32).

Run the server:
```
python -m excel_mcp.server
```
