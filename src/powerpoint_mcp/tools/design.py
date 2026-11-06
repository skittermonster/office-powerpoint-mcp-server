"""Theme application and institutional master footer formatting tools."""

from datetime import datetime
from typing import Annotated, Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import Field

from powerpoint_mcp.models import ThemeName
from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.utils.xml_utils import (
    hex_to_rgb,
    set_cell_border,
    set_cell_shading,
    set_slide_background,
)


def register_design_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register theme styling and master footer formatting tools with FastMCP."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Apply Theme",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def apply_theme(
        theme_name: Annotated[
            ThemeName,
            Field(
                description=("Theme palette: institutional_navy, modern_blue, corporate_gray, elegant_green, warm_red")
            ),
        ] = "institutional_navy",
        presentation_id: Annotated[
            str | None,
            Field(description="In-memory presentation ID"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="Path to presentation file within sandbox"),
        ] = None,
        slide_index: Annotated[
            int | None,
            Field(description="Specific 0-based slide index (applies to all slides if None)"),
        ] = None,
        apply_background: Annotated[
            bool,
            Field(description="Whether to apply theme background color via OpenXML"),
        ] = True,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Apply institutional color palette and typography across slides, tables, and cards."""
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            templates_data = manager.load_templates_data()
            color_schemes = templates_data.get("color_schemes", {})
            palette = color_schemes.get(theme_name, color_schemes.get("institutional_navy", {}))

            primary_color = palette.get("primary", "#1B365D")
            zebra_color = palette.get("zebra", "#F1F5F9")
            text_color = palette.get("text", "#1A1A1A")
            border_color = palette.get("border", "#CBD5E1")
            bg_color = palette.get("background", "#FFFFFF")
            fonts_cfg = palette.get("fonts", {})
            title_font = fonts_cfg.get("title", "Helvetica")
            body_font = fonts_cfg.get("body", "Arial")

            target_slides = (
                [prs.slides[slide_index]]
                if slide_index is not None and 0 <= slide_index < len(prs.slides)
                else list(prs.slides)
            )

            for slide in target_slides:
                if apply_background:
                    set_slide_background(slide, bg_color)

                for shape in slide.shapes:
                    # Update Tables
                    if shape.has_table:
                        tbl = shape.table
                        num_rows = len(tbl.rows)
                        num_cols = len(tbl.columns)
                        # Header
                        for ci in range(num_cols):
                            c = tbl.cell(0, ci)
                            set_cell_shading(c, primary_color)
                            set_cell_border(c, "bottom", color_hex=border_color, width_pt=1.5)
                            if c.text_frame and c.text_frame.paragraphs:
                                p = c.text_frame.paragraphs[0]
                                p.font.name = body_font
                                p.font.bold = True
                                p.font.color.rgb = hex_to_rgb("#FFFFFF")

                        # Body rows
                        for ri in range(1, num_rows):
                            row_bg = zebra_color if ri % 2 == 1 else "#FFFFFF"
                            for ci in range(num_cols):
                                c = tbl.cell(ri, ci)
                                set_cell_shading(c, row_bg)
                                set_cell_border(c, "bottom", color_hex=border_color, width_pt=0.75)
                                if c.text_frame and c.text_frame.paragraphs:
                                    p = c.text_frame.paragraphs[0]
                                    p.font.name = body_font
                                    p.font.color.rgb = hex_to_rgb(text_color)

                    # Update Text Frames
                    elif shape.has_text_frame and shape.text_frame:
                        is_title = shape == slide.shapes.title or "title" in shape.name.lower()
                        for p in shape.text_frame.paragraphs:
                            if is_title:
                                p.font.name = title_font
                                p.font.color.rgb = hex_to_rgb(primary_color)
                            else:
                                p.font.name = body_font
                                p.font.color.rgb = hex_to_rgb(text_color)

                    # Update Shapes / Cards
                    elif shape.shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = hex_to_rgb(palette.get("light", "#F1F5F9"))
                        shape.line.color.rgb = hex_to_rgb(border_color)

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "theme_applied": theme_name,
                "slides_updated": len(target_slides),
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Format Master Footer",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def format_master_footer(
        presentation_id: Annotated[
            str | None,
            Field(description="In-memory presentation ID"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="Path to presentation file within sandbox"),
        ] = None,
        footer_text: Annotated[
            str | None,
            Field(description="Confidentiality or disclaimer text on bottom-left"),
        ] = "Confidential - For Internal Use Only",
        include_slide_numbers: Annotated[
            bool,
            Field(description="Whether to include 'Slide X of Y' on bottom-right"),
        ] = True,
        include_date: Annotated[
            bool,
            Field(description="Whether to include current month/year in footer"),
        ] = True,
        date_text: Annotated[
            str | None,
            Field(description="Explicit date string override (e.g. 'November 2026')"),
        ] = None,
        skip_title_slide: Annotated[
            bool,
            Field(description="Whether to skip adding footer to slide 0"),
        ] = True,
        theme: Annotated[
            ThemeName,
            Field(description="Color scheme for footer accents"),
        ] = "institutional_navy",
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Add institutional divider line, slide numbering, confidentiality notice, and date across slides."""
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation
            total_slides = len(prs.slides)

            templates_data = manager.load_templates_data()
            color_schemes = templates_data.get("color_schemes", {})
            palette = color_schemes.get(theme, color_schemes.get("institutional_navy", {}))
            border_color = palette.get("border", "#CBD5E1")
            muted_color = palette.get("muted_text", "#64748B")
            body_font = palette.get("fonts", {}).get("body", "Arial")

            resolved_date = date_text or (datetime.now().strftime("%B %Y") if include_date else None)

            applied_slides = 0
            for idx, slide in enumerate(prs.slides):
                if idx == 0 and skip_title_slide:
                    continue

                # Footer divider line
                line = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(1.0),
                    Inches(6.85),
                    Inches(11.333),
                    Inches(0.015),
                )
                line.fill.solid()
                line.fill.fore_color.rgb = hex_to_rgb(border_color)
                line.line.fill.background()

                # Left confidentiality notice
                if footer_text:
                    left_tb = slide.shapes.add_textbox(
                        Inches(1.0),
                        Inches(6.88),
                        Inches(5.5),
                        Inches(0.35),
                    )
                    l_tf = left_tb.text_frame
                    l_tf.word_wrap = True
                    l_p = l_tf.paragraphs[0]
                    l_p.text = footer_text
                    l_p.font.name = body_font
                    l_p.font.size = Pt(8.5)
                    l_p.font.color.rgb = hex_to_rgb(muted_color)
                    l_p.alignment = PP_ALIGN.LEFT

                # Right metadata (Date | Slide X of Y)
                right_parts = []
                if resolved_date:
                    right_parts.append(resolved_date)
                if include_slide_numbers:
                    right_parts.append(f"Slide {idx + 1} of {total_slides}")

                if right_parts:
                    right_str = " | ".join(right_parts)
                    right_tb = slide.shapes.add_textbox(
                        Inches(6.833),
                        Inches(6.88),
                        Inches(5.5),
                        Inches(0.35),
                    )
                    r_tf = right_tb.text_frame
                    r_tf.word_wrap = True
                    r_p = r_tf.paragraphs[0]
                    r_p.text = right_str
                    r_p.font.name = body_font
                    r_p.font.size = Pt(8.5)
                    r_p.font.color.rgb = hex_to_rgb(muted_color)
                    r_p.alignment = PP_ALIGN.RIGHT

                applied_slides += 1

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "slides_formatted": applied_slides,
                "total_slides": total_slides,
                "footer_text": footer_text,
                "date_included": bool(resolved_date),
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}
