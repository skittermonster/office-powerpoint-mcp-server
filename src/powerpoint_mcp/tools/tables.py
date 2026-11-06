"""Institutional financial and data table tools for PowerPoint MCP Server."""

from typing import Annotated, Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import Field

from powerpoint_mcp.models import TableCellUpdate, ThemeName
from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.utils.geometry import calculate_column_widths
from powerpoint_mcp.utils.xml_utils import (
    hex_to_rgb,
    set_cell_border,
    set_cell_margins,
    set_cell_shading,
)


def _is_numeric(val: Any) -> bool:
    """Helper to detect if a table cell value represents numeric financial data."""
    if isinstance(val, (int, float)):
        return True
    s = str(val).strip().rstrip("%").lstrip("$€£¥").replace(",", "")
    try:
        float(s)
        return True
    except ValueError:
        return False


def register_table_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register financial and data table tools with FastMCP."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Data Table",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def add_data_table(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide to add table to"),
        ],
        headers: Annotated[
            list[str],
            Field(description="List of column header labels"),
        ],
        rows: Annotated[
            list[list[Any]],
            Field(description="2D array of row cell data"),
        ],
        presentation_id: Annotated[
            str | None,
            Field(description="In-memory presentation ID"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="Path to presentation file within sandbox"),
        ] = None,
        left: Annotated[
            float,
            Field(description="Left coordinate in inches"),
        ] = 1.0,
        top: Annotated[
            float,
            Field(description="Top coordinate in inches"),
        ] = 1.8,
        width: Annotated[
            float,
            Field(description="Total table width in inches"),
        ] = 11.333,
        height: Annotated[
            float | None,
            Field(description="Total table height in inches (auto-calculated if None)"),
        ] = None,
        column_widths: Annotated[
            list[float] | None,
            Field(description="Explicit column widths in inches (auto-calculated if None)"),
        ] = None,
        theme: Annotated[
            ThemeName,
            Field(description="Color scheme to use for table styling"),
        ] = "institutional_navy",
        footnote: Annotated[
            str | None,
            Field(description="Optional source or disclaimer note under the table"),
        ] = None,
        highlight_first_column: Annotated[
            bool,
            Field(description="Whether to style first column with bold font"),
        ] = False,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Create a publication-ready institutional data table on a slide.

        Features:
        - Automatic column width calculation with wider category/label column 0
        - Institutional styled header with theme primary background (#1B365D) and bold white text
        - Alternating row zebra striping with slate gray fill (#F1F5F9)
        - Subtle OpenXML horizontal bottom borders (#CBD5E1) on every row
        - Intelligent alignment: Left-aligned for text/labels, Right-aligned for numeric data
        - Clean cell margins / padding and optional source footnote
        """
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            if slide_index < 0 or slide_index >= len(prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range (0..{len(prs.slides) - 1})")

            slide = prs.slides[slide_index]
            templates_data = manager.load_templates_data()
            color_schemes = templates_data.get("color_schemes", {})
            palette = color_schemes.get(theme, color_schemes.get("institutional_navy", {}))

            primary_color = palette.get("primary", "#1B365D")
            zebra_color = palette.get("zebra", "#F1F5F9")
            text_color = palette.get("text", "#1A1A1A")
            border_color = palette.get("border", "#CBD5E1")
            fonts_cfg = palette.get("fonts", {})
            body_font = fonts_cfg.get("body", "Arial")

            num_cols = len(headers)
            num_rows = len(rows) + 1  # headers + data rows

            if num_cols == 0:
                raise ValueError("headers list cannot be empty")

            # Column widths
            if column_widths and len(column_widths) == num_cols:
                col_w_list = column_widths
            else:
                col_w_list = calculate_column_widths(width, headers, rows)

            # Row height estimation
            calculated_height = height or max(2.0, num_rows * 0.42)

            table_shape = slide.shapes.add_table(
                num_rows,
                num_cols,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(calculated_height),
            )
            table = table_shape.table

            # Assign column widths
            for c_idx, c_w in enumerate(col_w_list):
                table.columns[c_idx].width = Inches(c_w)

            # Style Header Row (row 0)
            for c_idx, header_text in enumerate(headers):
                cell = table.cell(0, c_idx)
                cell.text = str(header_text)
                set_cell_shading(cell, primary_color)
                set_cell_margins(cell, left_inches=0.12, right_inches=0.12, top_inches=0.08, bottom_inches=0.08)
                set_cell_border(cell, "bottom", color_hex=border_color, width_pt=1.5)

                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    p.font.name = body_font
                    p.font.bold = True
                    p.font.size = Pt(11)
                    p.font.color.rgb = hex_to_rgb("#FFFFFF")
                    p.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT

            # Style Data Rows
            for r_idx, row_data in enumerate(rows):
                table_r = r_idx + 1
                row_bg = zebra_color if r_idx % 2 == 1 else "#FFFFFF"

                for c_idx in range(num_cols):
                    cell = table.cell(table_r, c_idx)
                    val = row_data[c_idx] if c_idx < len(row_data) else ""
                    cell.text = str(val) if val is not None else ""

                    set_cell_shading(cell, row_bg)
                    set_cell_margins(cell, left_inches=0.12, right_inches=0.12, top_inches=0.07, bottom_inches=0.07)
                    set_cell_border(cell, "bottom", color_hex=border_color, width_pt=0.75)

                    if cell.text_frame and cell.text_frame.paragraphs:
                        p = cell.text_frame.paragraphs[0]
                        p.font.name = body_font
                        p.font.size = Pt(10)
                        p.font.color.rgb = hex_to_rgb(text_color)

                        if c_idx == 0 and highlight_first_column:
                            p.font.bold = True
                            p.alignment = PP_ALIGN.LEFT
                        elif c_idx == 0:
                            p.alignment = PP_ALIGN.LEFT
                        else:
                            p.alignment = PP_ALIGN.RIGHT if _is_numeric(val) else PP_ALIGN.LEFT

            # Add footnote if provided
            if footnote:
                fn_top = top + calculated_height + 0.08
                fn_box = slide.shapes.add_textbox(
                    Inches(left),
                    Inches(fn_top),
                    Inches(width),
                    Inches(0.35),
                )
                fn_tf = fn_box.text_frame
                fn_tf.word_wrap = True
                fn_p = fn_tf.paragraphs[0]
                fn_p.text = f"Source: {footnote}"
                fn_p.font.name = body_font
                fn_p.font.size = Pt(8.5)
                fn_p.font.italic = True
                fn_p.font.color.rgb = hex_to_rgb(palette.get("muted_text", "#64748B"))

            manager.auto_save_if_needed(record)

            shape_index = len(slide.shapes) - (2 if footnote else 1)
            return {
                "status": "success",
                "slide_index": slide_index,
                "shape_index": shape_index,
                "rows": num_rows,
                "columns": num_cols,
                "column_widths": col_w_list,
                "theme": theme,
                "has_footnote": bool(footnote),
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Update Table Cells",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def update_table_cells(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide containing the table"),
        ],
        shape_index: Annotated[
            int,
            Field(ge=0, description="0-based shape index of the table"),
        ],
        updates: Annotated[
            list[TableCellUpdate],
            Field(description="List of cell update specifications (row, col, text, colors, font)"),
        ],
        presentation_id: Annotated[
            str | None,
            Field(description="In-memory presentation ID"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="Path to presentation file within sandbox"),
        ] = None,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Update specific cells in an existing PowerPoint table.

        Allows modifying cell text, font size, bold/italic, text color,
        background shading, and text alignment.
        """
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            if slide_index < 0 or slide_index >= len(prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range")

            slide = prs.slides[slide_index]
            if shape_index < 0 or shape_index >= len(slide.shapes):
                raise IndexError(f"shape_index {shape_index} out of range")

            shape = slide.shapes[shape_index]
            if not shape.has_table:
                raise ValueError(f"Shape at index {shape_index} is not a table")

            table = shape.table
            applied_count = 0

            align_map = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }

            for u in updates:
                r, c = u.row, u.col
                if r < 0 or r >= len(table.rows) or c < 0 or c >= len(table.columns):
                    continue

                cell = table.cell(r, c)
                if u.text is not None:
                    cell.text = u.text

                if u.bg_color is not None:
                    set_cell_shading(cell, u.bg_color)

                if cell.text_frame and cell.text_frame.paragraphs:
                    p = cell.text_frame.paragraphs[0]
                    if u.font_size is not None:
                        p.font.size = Pt(u.font_size)
                    if u.bold is not None:
                        p.font.bold = u.bold
                    if u.italic is not None:
                        p.font.italic = u.italic
                    if u.text_color is not None:
                        p.font.color.rgb = hex_to_rgb(u.text_color)
                    if u.alignment and u.alignment in align_map:
                        p.alignment = align_map[u.alignment]

                applied_count += 1

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "slide_index": slide_index,
                "shape_index": shape_index,
                "cells_updated": applied_count,
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}
