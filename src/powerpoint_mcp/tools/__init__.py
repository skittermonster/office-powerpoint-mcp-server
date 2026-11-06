"""Tools package for Office PowerPoint MCP Server."""

from powerpoint_mcp.tools.content import register_content_tools
from powerpoint_mcp.tools.design import register_design_tools
from powerpoint_mcp.tools.lifecycle import register_lifecycle_tools
from powerpoint_mcp.tools.media_charts import register_media_chart_tools
from powerpoint_mcp.tools.slides import register_slide_tools
from powerpoint_mcp.tools.tables import register_table_tools

__all__ = [
    "register_content_tools",
    "register_design_tools",
    "register_lifecycle_tools",
    "register_media_chart_tools",
    "register_slide_tools",
    "register_table_tools",
]
