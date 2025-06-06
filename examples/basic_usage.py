"""Basic example demonstrating how to run the Excel MCP server."""

from excel_mcp.server import server

if __name__ == "__main__":
    # This will start the FastMCP server using streamable HTTP transport.
    # Connect to http://localhost:8000/ and call tools such as initialize_excel_link.
    server.run(transport="streamable-http")
