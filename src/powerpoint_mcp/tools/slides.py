"""Template-driven slide generation and slide management tools."""

import copy
from typing import Annotated, Any, Literal

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import Field

from powerpoint_mcp.models import TemplateName, ThemeName
from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.utils.geometry import calculate_column_widths, calculate_font_size
from powerpoint_mcp.utils.xml_utils import (
    hex_to_rgb,
    set_cell_border,
    set_cell_margins,
    set_cell_shading,
)


def _render_card_with_bullets(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    card_title: str,
    bullets: list[str | dict],
    palette: dict[str, Any],
    body_font: str,
) -> None:
    """Helper to render an institutional card container with title and bullet points."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb(palette.get("light", "#F1F5F9"))
    card.line.color.rgb = hex_to_rgb(palette.get("border", "#CBD5E1"))
    card.line.width = Pt(1.0)

    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)
    tf.margin_bottom = Inches(0.25)

    p_title = tf.paragraphs[0]
    p_title.text = card_title
    p_title.font.name = body_font
    p_title.font.bold = True
    p_title.font.size = Pt(15)
    p_title.font.color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))

    for b in bullets:
        p_b = tf.add_paragraph()
        p_b.font.name = body_font
        if isinstance(b, str):
            p_b.text = b
            p_b.level = 0
            p_b.font.size = Pt(11.5)
            p_b.font.color.rgb = hex_to_rgb(palette.get("text", "#1A1A1A"))
        elif isinstance(b, dict):
            p_b.text = str(b.get("text", ""))
            p_b.level = int(b.get("level", 0))
            p_b.font.bold = bool(b.get("bold", False))
            p_b.font.italic = bool(b.get("italic", False))
            p_b.font.size = Pt(float(b.get("size_pt", 11.5)))
            c = b.get("color") or palette.get("text", "#1A1A1A")
            p_b.font.color.rgb = hex_to_rgb(c)


def register_slide_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register slide management and template tools with FastMCP."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="List Templates",
            read_only_hint=True,
            destructive_hint=False,
            idempotent_hint=True,
        )
    )
    def list_templates() -> dict[str, Any]:
        """List all available pre-calculated slide layout templates and color schemes."""
        try:
            data = manager.load_templates_data()
            templates_info = {}
            for k, v in data.get("templates", {}).items():
                templates_info[k] = {
                    "name": v.get("name", k),
                    "description": v.get("description", ""),
                }
            return {
                "status": "success",
                "templates": templates_info,
                "color_schemes": list(data.get("color_schemes", {}).keys()),
                "default_dimensions": data.get("metadata", {}).get("slide_dimensions", {}),
            }
        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Slide From Template",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def add_slide_from_template(
        template_name: Annotated[
            TemplateName,
            Field(
                description=(
                    "Template ID: title_slide, executive_summary, two_column_comparison, "
                    "kpi_metrics_grid, table_with_commentary, chart_with_takeaways"
                )
            ),
        ],
        title: Annotated[
            str,
            Field(description="Slide main title"),
        ],
        subtitle: Annotated[
            str | None,
            Field(description="Slide contextual subtitle"),
        ] = None,
        elements_data: Annotated[
            dict[str, Any] | None,
            Field(description="Data dictionary for template placeholders (cards, metrics, tables, charts)"),
        ] = None,
        color_scheme: Annotated[
            ThemeName,
            Field(description="Institutional color palette name"),
        ] = "institutional_navy",
        slide_index: Annotated[
            int | None,
            Field(description="Position index to insert slide at (appends if None)"),
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
        """Build institutional presentation slides from pre-calculated coordinate templates.

        Supported Templates:
        - `title_slide`: Authoritative title slide with decorative accent bar and author/date
        - `executive_summary`: Two structured cards for key observations and recommendations
        - `two_column_comparison`: Side-by-side comparison cards (current vs target, options A/B)
        - `kpi_metrics_grid`: 4 high-impact metric cards with big numbers, labels, and YoY trends
        - `table_with_commentary`: Institutional data table on the left with commentary on the right
        - `chart_with_takeaways`: Native editable chart on the left with analyst takeaways on the right
        """
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
            new_slide = prs.slides.add_slide(blank_layout)

            # Move if slide_index provided
            if slide_index is not None and 0 <= slide_index < len(prs.slides) - 1:
                elem = prs.slides._sldIdLst[-1]
                prs.slides._sldIdLst.remove(elem)
                prs.slides._sldIdLst.insert(slide_index, elem)
                actual_idx = slide_index
            else:
                actual_idx = len(prs.slides) - 1

            templates_data = manager.load_templates_data()
            color_schemes = templates_data.get("color_schemes", {})
            palette = color_schemes.get(color_scheme, color_schemes.get("institutional_navy", {}))
            fonts_cfg = palette.get("fonts", {})
            title_font = fonts_cfg.get("title", "Helvetica")
            body_font = fonts_cfg.get("body", "Arial")
            data_dict = elements_data or {}

            # Helper for standard slide header
            def add_standard_header(header_title: str, header_sub: str | None) -> float:
                t_box = new_slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.333), Inches(0.75))
                t_box.name = "Title"
                t_tf = t_box.text_frame
                t_tf.word_wrap = True
                p_t = t_tf.paragraphs[0]
                p_t.text = header_title
                p_t.font.name = title_font
                p_t.font.bold = True
                font_sz = calculate_font_size(header_title, 11.333, 0.75, role="title", min_size=20, max_size=32)
                p_t.font.size = Pt(font_sz)
                p_t.font.color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))

                c_top = 1.45
                if header_sub:
                    s_box = new_slide.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(11.333), Inches(0.45))
                    s_tf = s_box.text_frame
                    s_tf.word_wrap = True
                    p_s = s_tf.paragraphs[0]
                    p_s.text = header_sub
                    p_s.font.name = body_font
                    p_s.font.size = Pt(13)
                    p_s.font.color.rgb = hex_to_rgb(palette.get("muted_text", "#64748B"))
                    c_top = 1.85

                line = new_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(1.0),
                    Inches(c_top),
                    Inches(11.333),
                    Inches(0.02),
                )
                line.fill.solid()
                line.fill.fore_color.rgb = hex_to_rgb(palette.get("border", "#CBD5E1"))
                line.line.fill.background()
                return c_top + 0.25

            # 1. title_slide
            if template_name == "title_slide":
                # Accent bar
                bar = new_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(1.5),
                    Inches(2.0),
                    Inches(1.2),
                    Inches(0.08),
                )
                bar.fill.solid()
                bar.fill.fore_color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))
                bar.line.fill.background()

                # Main Title
                title_box = new_slide.shapes.add_textbox(Inches(1.5), Inches(2.3), Inches(10.333), Inches(1.8))
                title_box.name = "Title"
                t_tf = title_box.text_frame
                t_tf.word_wrap = True
                p0 = t_tf.paragraphs[0]
                p0.text = title
                p0.font.name = title_font
                p0.font.bold = True
                p0.font.size = Pt(calculate_font_size(title, 10.333, 1.8, role="title", min_size=24, max_size=42))
                p0.font.color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))

                # Subtitle
                sub_text = subtitle or data_dict.get("subtitle")
                if sub_text:
                    sub_box = new_slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(10.333), Inches(1.0))
                    s_tf = sub_box.text_frame
                    s_tf.word_wrap = True
                    p_sub = s_tf.paragraphs[0]
                    p_sub.text = sub_text
                    p_sub.font.name = body_font
                    p_sub.font.size = Pt(20)
                    p_sub.font.color.rgb = hex_to_rgb(palette.get("muted_text", "#64748B"))

                # Author / Org / Date
                author_text = data_dict.get("author") or data_dict.get("organization") or ""
                date_text = data_dict.get("date") or ""
                footer_str = " | ".join(filter(None, [author_text, date_text]))
                if footer_str:
                    a_box = new_slide.shapes.add_textbox(Inches(1.5), Inches(5.5), Inches(10.333), Inches(0.8))
                    a_tf = a_box.text_frame
                    a_tf.word_wrap = True
                    p_a = a_tf.paragraphs[0]
                    p_a.text = footer_str
                    p_a.font.name = body_font
                    p_a.font.bold = True
                    p_a.font.size = Pt(13)
                    p_a.font.color.rgb = hex_to_rgb(palette.get("secondary", "#0F172A"))

            # 2. executive_summary
            elif template_name == "executive_summary":
                content_top = add_standard_header(title, subtitle)
                c_height = 7.0 - content_top
                _render_card_with_bullets(
                    new_slide,
                    left=1.0,
                    top=content_top,
                    width=5.45,
                    height=c_height,
                    card_title=data_dict.get("left_title", "Key Observations & Findings"),
                    bullets=data_dict.get("left_bullets", ["Primary observation on business operations."]),
                    palette=palette,
                    body_font=body_font,
                )
                _render_card_with_bullets(
                    new_slide,
                    left=6.883,
                    top=content_top,
                    width=5.45,
                    height=c_height,
                    card_title=data_dict.get("right_title", "Strategic Recommendations"),
                    bullets=data_dict.get("right_bullets", ["Recommended capital and strategic allocations."]),
                    palette=palette,
                    body_font=body_font,
                )

            # 3. two_column_comparison
            elif template_name == "two_column_comparison":
                content_top = add_standard_header(title, subtitle)
                c_height = 7.0 - content_top
                _render_card_with_bullets(
                    new_slide,
                    left=1.0,
                    top=content_top,
                    width=5.45,
                    height=c_height,
                    card_title=data_dict.get("col1_title", "Option A / Current State"),
                    bullets=data_dict.get("col1_bullets", ["Baseline operational architecture."]),
                    palette=palette,
                    body_font=body_font,
                )
                _render_card_with_bullets(
                    new_slide,
                    left=6.883,
                    top=content_top,
                    width=5.45,
                    height=c_height,
                    card_title=data_dict.get("col2_title", "Option B / Target State"),
                    bullets=data_dict.get("col2_bullets", ["Proposed modernized target capabilities."]),
                    palette=palette,
                    body_font=body_font,
                )

            # 4. kpi_metrics_grid
            elif template_name == "kpi_metrics_grid":
                content_top = add_standard_header(title, subtitle)
                metrics = data_dict.get(
                    "metrics",
                    [
                        {"value": "$45.2M", "label": "Revenue", "subtext": "+28% YoY"},
                        {"value": "68.5%", "label": "Gross Margin", "subtext": "+340 bps"},
                        {"value": "124%", "label": "Net Retention", "subtext": "Top Quartile"},
                        {"value": "2.4x", "label": "LTV / CAC", "subtext": "Payback 11 mos"},
                    ],
                )
                c_height = 7.0 - content_top
                c_width = 2.65
                gap = 0.24
                for m_idx, m in enumerate(metrics[:4]):
                    m_left = 1.0 + m_idx * (c_width + gap)
                    card = new_slide.shapes.add_shape(
                        MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(m_left),
                        Inches(content_top),
                        Inches(c_width),
                        Inches(c_height),
                    )
                    card.fill.solid()
                    card.fill.fore_color.rgb = hex_to_rgb(palette.get("light", "#F1F5F9"))
                    card.line.color.rgb = hex_to_rgb(palette.get("border", "#CBD5E1"))
                    card.line.width = Pt(1.0)

                    tf = card.text_frame
                    tf.word_wrap = True
                    tf.margin_left = Inches(0.2)
                    tf.margin_right = Inches(0.2)
                    tf.margin_top = Inches(0.4)

                    val_str = str(m.get("value", ""))
                    p_val = tf.paragraphs[0]
                    p_val.text = val_str
                    p_val.font.name = title_font
                    p_val.font.bold = True
                    kpi_font_sz = calculate_font_size(val_str, c_width, 1.2, role="kpi", min_size=24, max_size=40)
                    p_val.font.size = Pt(kpi_font_sz)
                    p_val.font.color.rgb = hex_to_rgb(palette.get("primary", "#1B365D"))

                    p_lbl = tf.add_paragraph()
                    p_lbl.text = str(m.get("label", ""))
                    p_lbl.font.name = body_font
                    p_lbl.font.bold = True
                    p_lbl.font.size = Pt(12)
                    p_lbl.font.color.rgb = hex_to_rgb(palette.get("text", "#1A1A1A"))

                    if m.get("subtext"):
                        p_sub = tf.add_paragraph()
                        p_sub.text = str(m.get("subtext"))
                        p_sub.font.name = body_font
                        p_sub.font.size = Pt(11)
                        p_sub.font.color.rgb = hex_to_rgb(palette.get("accent1", "#2563EB"))

            # 5. table_with_commentary
            elif template_name == "table_with_commentary":
                content_top = add_standard_header(title, subtitle)
                c_height = 7.0 - content_top
                tbl_w = 7.6
                headers = data_dict.get("table_headers", ["Segment", "Revenue ($M)", "Growth"])
                rows = data_dict.get(
                    "table_rows",
                    [
                        ["Enterprise Cloud", "$84.2M", "+32%"],
                        ["Commercial Software", "$41.0M", "+18%"],
                        ["Professional Services", "$12.8M", "+6%"],
                    ],
                )

                num_rows = len(rows) + 1
                num_cols = len(headers)
                col_widths = calculate_column_widths(tbl_w, headers, rows)
                tbl_shape = new_slide.shapes.add_table(
                    num_rows,
                    num_cols,
                    Inches(1.0),
                    Inches(content_top),
                    Inches(tbl_w),
                    Inches(min(c_height, num_rows * 0.45)),
                )
                tbl = tbl_shape.table
                for ci, cw in enumerate(col_widths):
                    tbl.columns[ci].width = Inches(cw)

                for ci, h in enumerate(headers):
                    cell = tbl.cell(0, ci)
                    cell.text = str(h)
                    set_cell_shading(cell, palette.get("primary", "#1B365D"))
                    set_cell_margins(cell, 0.1, 0.1, 0.08, 0.08)
                    set_cell_border(cell, "bottom", color_hex=palette.get("border", "#CBD5E1"), width_pt=1.5)
                    p = cell.text_frame.paragraphs[0]
                    p.font.name = body_font
                    p.font.bold = True
                    p.font.size = Pt(10.5)
                    p.font.color.rgb = hex_to_rgb("#FFFFFF")
                    p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT

                for ri, row in enumerate(rows):
                    row_bg = palette.get("zebra", "#F1F5F9") if ri % 2 == 1 else "#FFFFFF"
                    for ci in range(num_cols):
                        cell = tbl.cell(ri + 1, ci)
                        val = row[ci] if ci < len(row) else ""
                        cell.text = str(val)
                        set_cell_shading(cell, row_bg)
                        set_cell_margins(cell, 0.1, 0.1, 0.07, 0.07)
                        set_cell_border(cell, "bottom", color_hex=palette.get("border", "#CBD5E1"), width_pt=0.75)
                        p = cell.text_frame.paragraphs[0]
                        p.font.name = body_font
                        p.font.size = Pt(10)
                        p.font.color.rgb = hex_to_rgb(palette.get("text", "#1A1A1A"))
                        p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT

                # Right commentary card
                _render_card_with_bullets(
                    new_slide,
                    left=8.9,
                    top=content_top,
                    width=3.433,
                    height=c_height,
                    card_title=data_dict.get("commentary_title", "Key Takeaways"),
                    bullets=data_dict.get("commentary_bullets", ["Enterprise Cloud driving margin expansion."]),
                    palette=palette,
                    body_font=body_font,
                )

            # 6. chart_with_takeaways
            elif template_name == "chart_with_takeaways":
                content_top = add_standard_header(title, subtitle)
                c_height = 7.0 - content_top
                chart_w = 7.6
                c_type_str = data_dict.get("chart_type", "COLUMN_CLUSTERED")
                cats = data_dict.get("categories", ["Q1", "Q2", "Q3", "Q4"])
                series_data = data_dict.get("series", [{"name": "2026", "values": [25, 30, 38, 45]}])

                chart_data = CategoryChartData()
                chart_data.categories = cats
                for s in series_data:
                    chart_data.add_series(s.get("name", "Series"), tuple(s.get("values", [])))

                chart_map = {
                    "COLUMN_CLUSTERED": XL_CHART_TYPE.COLUMN_CLUSTERED,
                    "LINE": XL_CHART_TYPE.LINE,
                    "PIE": XL_CHART_TYPE.PIE,
                    "BAR_CLUSTERED": XL_CHART_TYPE.BAR_CLUSTERED,
                }
                c_type = chart_map.get(c_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)
                chart_shape = new_slide.shapes.add_chart(
                    c_type,
                    Inches(1.0),
                    Inches(content_top),
                    Inches(chart_w),
                    Inches(c_height),
                    chart_data,
                )
                chart = chart_shape.chart
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
                if c_type_str != "PIE":
                    chart.value_axis.has_major_gridlines = True

                # Right takeaways card
                _render_card_with_bullets(
                    new_slide,
                    left=8.9,
                    top=content_top,
                    width=3.433,
                    height=c_height,
                    card_title=data_dict.get("takeaways_title", "Performance Highlights"),
                    bullets=data_dict.get("takeaways_bullets", ["Accelerating momentum through Q4."]),
                    palette=palette,
                    body_font=body_font,
                )

            else:
                raise ValueError(f"Unknown template_name '{template_name}'")

            manager.auto_save_if_needed(record)

            return {
                "status": "success",
                "template_name": template_name,
                "slide_index": actual_idx,
                "color_scheme": color_scheme,
                "slide_count": len(prs.slides),
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Slides",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def manage_slides(
        action: Annotated[
            Literal["add", "duplicate", "delete", "reorder", "populate_placeholder"],
            Field(description="Slide structural action: add, duplicate, delete, reorder, populate_placeholder"),
        ],
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of target slide"),
        ] = 0,
        to_index: Annotated[
            int | None,
            Field(description="Target destination index (required for 'reorder')"),
        ] = None,
        layout_index: Annotated[
            int,
            Field(description="Layout index for 'add' action (default 6 for blank layout)"),
        ] = 6,
        placeholder_index: Annotated[
            int | None,
            Field(description="Placeholder index for 'populate_placeholder'"),
        ] = None,
        placeholder_name: Annotated[
            str | None,
            Field(description="Placeholder name / title for 'populate_placeholder'"),
        ] = None,
        text: Annotated[
            str | None,
            Field(description="Text content to insert into placeholder"),
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
        """Perform slide structural operations: duplicate, delete, reorder, or populate placeholders."""
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            if action == "add":
                valid_layout_idx = layout_index if 0 <= layout_index < len(prs.slide_layouts) else 0
                new_slide = prs.slides.add_slide(prs.slide_layouts[valid_layout_idx])
                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "add",
                    "slide_index": len(prs.slides) - 1,
                    "slide_count": len(prs.slides),
                }

            elif action == "delete":
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise IndexError(f"slide_index {slide_index} out of range")
                rId = prs.slides._sldIdLst[slide_index].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[slide_index]
                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "delete",
                    "deleted_index": slide_index,
                    "slide_count": len(prs.slides),
                }

            elif action == "duplicate":
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise IndexError(f"slide_index {slide_index} out of range")
                source_slide = prs.slides[slide_index]
                new_slide = prs.slides.add_slide(source_slide.slide_layout)
                for shape in source_slide.shapes:
                    new_slide.shapes._spTree.append(copy.deepcopy(shape._element))
                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "duplicate",
                    "source_index": slide_index,
                    "new_slide_index": len(prs.slides) - 1,
                    "slide_count": len(prs.slides),
                }

            elif action == "reorder":
                if to_index is None:
                    raise ValueError("to_index is required for 'reorder' action")
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise IndexError(f"slide_index {slide_index} out of range")
                if to_index < 0 or to_index >= len(prs.slides):
                    raise IndexError(f"to_index {to_index} out of range")
                elem = prs.slides._sldIdLst[slide_index]
                prs.slides._sldIdLst.remove(elem)
                prs.slides._sldIdLst.insert(to_index, elem)
                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "reorder",
                    "from_index": slide_index,
                    "to_index": to_index,
                    "slide_count": len(prs.slides),
                }

            elif action == "populate_placeholder":
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise IndexError(f"slide_index {slide_index} out of range")
                slide = prs.slides[slide_index]
                target_ph = None

                if placeholder_index is not None:
                    try:
                        target_ph = slide.placeholders[placeholder_index]
                    except KeyError:
                        pass

                if target_ph is None and placeholder_name:
                    for ph in slide.placeholders:
                        if placeholder_name.lower() in ph.name.lower():
                            target_ph = ph
                            break

                if target_ph is None:
                    raise ValueError(
                        f"Placeholder not found on slide {slide_index}. "
                        f"Available: {[p.name for p in slide.placeholders]}"
                    )

                if text is not None:
                    target_ph.text = text

                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "populate_placeholder",
                    "slide_index": slide_index,
                    "placeholder_name": target_ph.name,
                }

            else:
                raise ValueError(f"Unknown action '{action}'")

        except Exception as e:
            return {"status": "error", "error": str(e)}
