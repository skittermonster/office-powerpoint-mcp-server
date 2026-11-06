"""Office PowerPoint MCP Server entry point and configuration."""

import argparse
import sys
from pathlib import Path
from typing import Sequence

from fastmcp import FastMCP

from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.tools.content import register_content_tools
from powerpoint_mcp.tools.design import register_design_tools
from powerpoint_mcp.tools.lifecycle import register_lifecycle_tools
from powerpoint_mcp.tools.media_charts import register_media_chart_tools
from powerpoint_mcp.tools.slides import register_slide_tools
from powerpoint_mcp.tools.tables import register_table_tools


def create_server(manager: SessionManager = session_manager) -> FastMCP:
    """Factory to instantiate and configure the FastMCP server with all tools."""
    server = FastMCP(
        name="office-powerpoint-mcp-server",
    )

    # Register all consolidated tool modules
    register_lifecycle_tools(server, manager)
    register_slide_tools(server, manager)
    register_content_tools(server, manager)
    register_table_tools(server, manager)
    register_media_chart_tools(server, manager)
    register_design_tools(server, manager)

    return server


# Default server instance
mcp = create_server(session_manager)


def parse_args(args: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse CLI arguments supporting transport and sandboxed directories."""
    parser = argparse.ArgumentParser(
        prog="ppt-mcp-server",
        description="Office PowerPoint MCP Server with path sandboxing and OpenXML styling",
    )
    parser.add_argument(
        "transport",
        nargs="?",
        default="stdio",
        choices=["stdio", "http", "sse", "streamable-http"],
        help="MCP transport protocol (default: stdio)",
    )
    parser.add_argument(
        "--allow-dir",
        action="append",
        dest="allow_dirs",
        metavar="PATH",
        help="Allowed directory for PowerPoint decks (can be specified multiple times; defaults to cwd)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host for HTTP transport (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transport (default: 8000)",
    )
    return parser.parse_args(args)


def main() -> None:
    """CLI entry point for ppt-mcp-server."""
    parsed = parse_args(sys.argv[1:])

    # Configure sandboxed directories
    if parsed.allow_dirs:
        session_manager.set_allowed_dirs(parsed.allow_dirs)
    else:
        session_manager.set_allowed_dirs([Path.cwd()])

    if parsed.transport == "stdio":
        mcp.run(transport="stdio")
    else:
        mcp.run(transport=parsed.transport, host=parsed.host, port=parsed.port)


if __name__ == "__main__":
    main()
