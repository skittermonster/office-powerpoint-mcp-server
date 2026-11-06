"""Utilities package for Office PowerPoint MCP Server."""

from powerpoint_mcp.utils.geometry import (
    calculate_column_widths,
    calculate_font_size,
    fit_image_in_bounding_box,
)
from powerpoint_mcp.utils.xml_utils import (
    hex_to_rgb,
    normalize_hex_color,
    set_cell_border,
    set_cell_borders,
    set_cell_margins,
    set_cell_shading,
    set_slide_background,
)

__all__ = [
    "calculate_column_widths",
    "calculate_font_size",
    "fit_image_in_bounding_box",
    "hex_to_rgb",
    "normalize_hex_color",
    "set_cell_border",
    "set_cell_borders",
    "set_cell_margins",
    "set_cell_shading",
    "set_slide_background",
]
