"""Content manipulation tools: text boxes, bulleted slides, and shapes."""

from typing import Annotated, Any, Literal

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import Field

from powerpoint_mcp.models import (
    BulletItem,
    SupportedShapeType,
    TextAlignment,
    ThemeName,
)
from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.utils.geometry import calculate_font_size
from powerpoint_mcp.utils.xml_utils import hex_to_rgb


def register_content_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register text, bullet slide, and shape tools with FastMCP."""

    align_map = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Text",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def manage_text(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide"),
        ],
        action: Annotated[
            Literal["add", "update", "clear"],
            Field(description="Text action: 'add' (creates text box), 'update' (modifies shape), 'clear'"),
        ],
        text: Annotated[
            str | None,
            Field(description="Plain text content or single paragraph"),
        ] = None,
        shape_index: Annotated[
            int | None,
            Field(description="Shape index to update or clear (required for update/clear)"),
        ] = None,
        left: Annotated[
            float,
            Field(description="Left coordinate in inches (for 'add')"),
        ] = 1.0,
        top: Annotated[
            float,
            Field(description="Top coordinate in inches (for 'add')"),
        ] = 1.0,
        width: Annotated[
            float,
            Field(description="Width in inches (for 'add')"),
        ] = 6.0,
        height: Annotated[
            float,
            Field(description="Height in inches (for 'add')"),
        ] = 2.0,
        font_size: Annotated[
            float | None,
            Field(description="Font size in points (auto-fit if omitted and auto_fit=True)"),
        ] = None,
        font_name: Annotated[
            str | None,
            Field(description="Font family name (e.g. 'Helvetica', 'Arial')"),
        ] = "Arial",
        font_color: Annotated[
            str | None,
            Field(description="Hex font color e.g. '#1A1A1A'"),
        ] = "#1A1A1A",
        bold: Annotated[
            bool,
            Field(description="Whether text is bold"),
        ] = False,
        italic: Annotated[
            bool,
            Field(description="Whether text is italic"),
        ] = False,
        alignment: Annotated[
            TextAlignment,
            Field(description="Text horizontal alignment: left, center, right, justify"),
        ] = "left",
        bullet_points: Annotated[
            list[BulletItem | str] | None,
            Field(description="Multi-level bullet points or nested items"),
        ] = None,
        auto_fit: Annotated[
            bool,
            Field(description="Automatically scale down font size based on character count"),
        ] = True,
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
        """Add, update, or clear text boxes and multi-level bulleted lists.

        Features:
        - Multi-level indentation (levels 0-4)
        - Run-level bold, color, and size overrides
        - Automatic shrink-to-fit font scaling preventing box overflow
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

            if action == "add":
                textbox = slide.shapes.add_textbox(
                    Inches(left),
                    Inches(top),
                    Inches(width),
                    Inches(height),
                )
                tf = textbox.text_frame
                tf.word_wrap = True
                target_shape_idx = len(slide.shapes) - 1
            elif action in ("update", "clear"):
                if shape_index is None or shape_index < 0 or shape_index >= len(slide.shapes):
                    raise IndexError(f"Valid shape_index required for action '{action}'")
                shape = slide.shapes[shape_index]
                if not shape.has_text_frame:
                    raise ValueError(f"Shape {shape_index} does not have a text frame")
                tf = shape.text_frame
                tf.word_wrap = True
                target_shape_idx = shape_index
                if action == "clear":
                    tf.clear()
                    manager.auto_save_if_needed(record)
                    return {"status": "success", "action": "clear", "shape_index": shape_index}
            else:
                raise ValueError(f"Unknown action '{action}'")

            # Determine optimal font size
            full_text = (text or "") + (" ".join(str(b) for b in (bullet_points or [])))
            calculated_size = font_size
            if not calculated_size and auto_fit:
                calculated_size = calculate_font_size(
                    full_text,
                    width,
                    height,
                    role="body",
                    min_size=8.0,
                    max_size=28.0,
                )
            size_pt = calculated_size or 12.0

            # If plain text provided
            if text is not None and not bullet_points:
                tf.text = text
                for p in tf.paragraphs:
                    p.font.name = font_name
                    p.font.size = Pt(size_pt)
                    p.font.bold = bold
                    p.font.italic = italic
                    if font_color:
                        p.font.color.rgb = hex_to_rgb(font_color)
                    p.alignment = align_map.get(alignment, PP_ALIGN.LEFT)

            # If bullet points provided
            if bullet_points:
                tf.clear()
                for idx, item in enumerate(bullet_points):
                    p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                    if isinstance(item, str):
                        p.text = item
                        p.level = 0
                        p.font.bold = bold
                        p.font.italic = italic
                        p.font.size = Pt(size_pt)
                        if font_color:
                            p.font.color.rgb = hex_to_rgb(font_color)
                    elif isinstance(item, (dict, BulletItem)):
                        b_obj = item if isinstance(item, BulletItem) else BulletItem(**item)
                        p.text = b_obj.text
                        p.level = b_obj.level
                        p.font.bold = b_obj.bold
                        p.font.italic = b_obj.italic
                        p.font.size = Pt(b_obj.size_pt or size_pt)
                        c = b_obj.color or font_color
                        if c:
                            p.font.color.rgb = hex_to_rgb(c)

                    p.font.name = font_name
                    p.alignment = align_map.get(alignment, PP_ALIGN.LEFT)

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "action": action,
                "slide_index": slide_index,
                "shape_index": target_shape_idx,
                "font_size_pt": size_pt,
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Bullet Slide",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def add_bullet_slide(
        title: Annotated[
            str,
            Field(description="Slide headline title"),
        ],
        bullet_points: Annotated[
            list[BulletItem | str],
            Field(description="List of strings or structured BulletItem with levels and formatting"),
        ],
        subtitle: Annotated[
            str | None,
            Field(description="Optional contextual subtitle under the header"),
        ] = None,
        slide_index: Annotated[
            int | None,
            Field(description="Position index to insert slide at (appends to end if None)"),
        ] = None,
        theme: Annotated[
            ThemeName,
            Field(description="Color scheme to use for typography and accents"),
        ] = "institutional_navy",
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
        """Quickly add an institutional structured bullet slide with header and divider."""
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
            new_slide = prs.slides.add_slide(blank_layout)

            # Move slide if slide_index requested
            if slide_index is not None and 0 <= slide_index < len(prs.slides) - 1:
                elem = prs.slides._sldIdLst[-1]
                prs.slides._sldIdLst.remove(elem)
                prs.slides._sldIdLst.insert(slide_index, elem)
                actual_idx = slide_index
            else:
                actual_idx = len(prs.slides) - 1

            templates_data = manager.load_templates_data()
            color_schemes = templates_data.get("color_schemes", {})
            palette = color_schemes.get(theme, color_schemes.get("institutional_navy", {}))
            fonts_cfg = palette.get("fonts", {})
            title_font = fonts_cfg.get("title", "Helvetica")
            body_font = fonts_cfg.get("body", "Arial")

            # Title Box
            title_box = new_slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(0.8))
            title_box.name = "Title"
            t_tf = title_box.text_frame
            t_tf.word_wrap = True
            tp = t_tf.paragraphs[0]
            tp.text = title
            tp.font.name = title_font
            tp.font.bold = True
            tp.font.size = Pt(28)
            tp.font.color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))

            content_top = 1.6
            if subtitle:
                sub_box = new_slide.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.333), Inches(0.45))
                s_tf = sub_box.text_frame
                s_tf.word_wrap = True
                sp = s_tf.paragraphs[0]
                sp.text = subtitle
                sp.font.name = body_font
                sp.font.size = Pt(13)
                sp.font.color.rgb = hex_to_rgb(palette.get("muted_text", "#64748B"))
                content_top = 1.9

            # Divider line
            line_shape = new_slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(1.0),
                Inches(content_top),
                Inches(11.333),
                Inches(0.02),
            )
            line_shape.fill.solid()
            line_shape.fill.fore_color.rgb = hex_to_rgb(palette.get("border", "#CBD5E1"))
            line_shape.line.fill.background()

            # Bullets Container
            body_top = content_top + 0.3
            body_height = 7.0 - body_top
            bullet_box = new_slide.shapes.add_textbox(
                Inches(1.0),
                Inches(body_top),
                Inches(11.333),
                Inches(body_height),
            )
            b_tf = bullet_box.text_frame
            b_tf.word_wrap = True

            for idx, item in enumerate(bullet_points):
                p = b_tf.add_paragraph() if idx > 0 else b_tf.paragraphs[0]
                if isinstance(item, str):
                    p.text = item
                    p.level = 0
                    p.font.size = Pt(14)
                    p.font.color.rgb = hex_to_rgb(palette.get("text", "#1A1A1A"))
                else:
                    b_obj = item if isinstance(item, BulletItem) else BulletItem(**item)
                    p.text = b_obj.text
                    p.level = b_obj.level
                    p.font.bold = b_obj.bold
                    p.font.italic = b_obj.italic
                    p.font.size = Pt(b_obj.size_pt or (14 - b_obj.level * 1.5))
                    c = b_obj.color or palette.get("text", "#1A1A1A")
                    p.font.color.rgb = hex_to_rgb(c)
                p.font.name = body_font

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "slide_index": actual_idx,
                "bullet_count": len(bullet_points),
                "theme": theme,
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Shapes",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def manage_shapes(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide"),
        ],
        action: Annotated[
            Literal["add", "delete"],
            Field(description="Shape action: 'add' or 'delete'"),
        ],
        shape_type: Annotated[
            SupportedShapeType,
            Field(description="Shape type: rectangle, rounded_rectangle, card, callout, arrow, oval"),
        ] = "rectangle",
        left: Annotated[
            float,
            Field(description="Left coordinate in inches"),
        ] = 1.0,
        top: Annotated[
            float,
            Field(description="Top coordinate in inches"),
        ] = 1.0,
        width: Annotated[
            float,
            Field(description="Width in inches"),
        ] = 4.0,
        height: Annotated[
            float,
            Field(description="Height in inches"),
        ] = 2.0,
        fill_color: Annotated[
            str | None,
            Field(description="Hex fill color (e.g. '#F1F5F9' for cards)"),
        ] = None,
        border_color: Annotated[
            str | None,
            Field(description="Hex border line color (e.g. '#CBD5E1')"),
        ] = None,
        border_width_pt: Annotated[
            float,
            Field(description="Border width in points"),
        ] = 1.0,
        text: Annotated[
            str | None,
            Field(description="Optional text inside the shape"),
        ] = None,
        text_color: Annotated[
            str | None,
            Field(description="Hex text color e.g. '#1A1A1A'"),
        ] = None,
        font_size: Annotated[
            float | None,
            Field(description="Font size in points"),
        ] = None,
        bold: Annotated[
            bool,
            Field(description="Whether inside text is bold"),
        ] = False,
        shape_index: Annotated[
            int | None,
            Field(description="Index of shape to delete (required for 'delete')"),
        ] = None,
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
        """Add styled visual cards, callout containers, arrows, or geometric shapes."""
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

            if action == "delete":
                if shape_index is None or shape_index < 0 or shape_index >= len(slide.shapes):
                    raise IndexError("Valid shape_index required to delete shape")
                target_shape = slide.shapes[shape_index]
                sp_elem = target_shape._element
                sp_elem.getparent().remove(sp_elem)
                manager.auto_save_if_needed(record)
                return {"status": "success", "action": "delete", "shape_index": shape_index}

            elif action == "add":
                shape_map = {
                    "rectangle": MSO_SHAPE.RECTANGLE,
                    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
                    "card": MSO_SHAPE.ROUNDED_RECTANGLE,
                    "callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
                    "arrow": MSO_SHAPE.RIGHT_ARROW,
                    "oval": MSO_SHAPE.OVAL,
                }
                mso_type = shape_map.get(shape_type, MSO_SHAPE.RECTANGLE)

                new_shape = slide.shapes.add_shape(
                    mso_type,
                    Inches(left),
                    Inches(top),
                    Inches(width),
                    Inches(height),
                )

                # Fill
                if fill_color:
                    new_shape.fill.solid()
                    new_shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
                else:
                    new_shape.fill.background()

                # Border line
                if border_color and border_width_pt > 0:
                    new_shape.line.color.rgb = hex_to_rgb(border_color)
                    new_shape.line.width = Pt(border_width_pt)
                else:
                    new_shape.line.fill.background()

                # Text inside shape
                if text is not None:
                    tf = new_shape.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = text
                    p.font.bold = bold
                    if font_size:
                        p.font.size = Pt(font_size)
                    if text_color:
                        p.font.color.rgb = hex_to_rgb(text_color)

                manager.auto_save_if_needed(record)

                target_shape_idx = len(slide.shapes) - 1
                return {
                    "status": "success",
                    "action": "add",
                    "shape_type": shape_type,
                    "slide_index": slide_index,
                    "shape_index": target_shape_idx,
                }

            else:
                raise ValueError(f"Unknown action '{action}'")

        except Exception as e:
            return {"status": "error", "error": str(e)}
