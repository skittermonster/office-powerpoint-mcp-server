"""Presentation lifecycle and inspection tools."""

from typing import Annotated, Any, Literal

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx.util import Inches
from pydantic import Field

from powerpoint_mcp.state import SessionManager, session_manager


def register_lifecycle_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register presentation lifecycle and inspection tools with FastMCP."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Manage Presentation",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def manage_presentation(
        action: Annotated[
            Literal["create", "open", "save", "close", "set_properties"],
            Field(description="Lifecycle action to perform on presentation"),
        ],
        presentation_id: Annotated[
            str | None,
            Field(description="Unique ID for in-memory presentation session"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="File path within allowed directory sandbox"),
        ] = None,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically persist changes to file_path on disk"),
        ] = True,
        width_inches: Annotated[
            float,
            Field(description="Slide width in inches (default: 13.333 for widescreen 16:9)"),
        ] = 13.333,
        height_inches: Annotated[
            float,
            Field(description="Slide height in inches (default: 7.5 for widescreen 16:9)"),
        ] = 7.5,
        title: Annotated[
            str | None,
            Field(description="Document core title property"),
        ] = None,
        author: Annotated[
            str | None,
            Field(description="Document core author / creator property"),
        ] = None,
        comments: Annotated[
            str | None,
            Field(description="Document core comments / description property"),
        ] = None,
    ) -> dict[str, Any]:
        """Create, open, save, close, or configure PowerPoint presentations.

        Defaults to widescreen 16:9 format (13.333 x 7.5 inches).
        Supports hybrid in-memory sessions and direct disk auto-saving.
        """
        try:
            if action == "create":
                record = manager.create_presentation(
                    presentation_id=presentation_id,
                    file_path=file_path,
                    width_inches=width_inches,
                    height_inches=height_inches,
                    auto_save=auto_save,
                )
                if title or author or comments:
                    props = record.presentation.core_properties
                    if title:
                        props.title = title
                    if author:
                        props.author = author
                    if comments:
                        props.comments = comments
                    manager.auto_save_if_needed(record)

                return {
                    "status": "success",
                    "action": "create",
                    "presentation_id": record.presentation_id,
                    "file_path": str(record.file_path) if record.file_path else None,
                    "auto_save": record.auto_save,
                    "slide_count": len(record.presentation.slides),
                    "dimensions": {
                        "width_inches": width_inches,
                        "height_inches": height_inches,
                        "aspect_ratio": "16:9",
                    },
                }

            elif action == "open":
                if not file_path:
                    raise ValueError("file_path is required for 'open' action")
                record = manager.open_presentation(
                    file_path=file_path,
                    presentation_id=presentation_id,
                    auto_save=auto_save,
                )
                return {
                    "status": "success",
                    "action": "open",
                    "presentation_id": record.presentation_id,
                    "file_path": str(record.file_path),
                    "slide_count": len(record.presentation.slides),
                    "auto_save": record.auto_save,
                }

            elif action == "save":
                record = manager.get_or_load_presentation(
                    presentation_id=presentation_id,
                    file_path=file_path,
                )
                saved_path = manager.save_record(record, target_path=file_path)
                return {
                    "status": "success",
                    "action": "save",
                    "presentation_id": record.presentation_id,
                    "file_path": str(saved_path),
                    "slide_count": len(record.presentation.slides),
                }

            elif action == "close":
                pid = presentation_id or manager.active_presentation_id
                if not pid:
                    raise ValueError("No active presentation to close")
                success = manager.close_presentation(pid)
                return {
                    "status": "success" if success else "not_found",
                    "action": "close",
                    "presentation_id": pid,
                    "active_presentation_id": manager.active_presentation_id,
                }

            elif action == "set_properties":
                record = manager.get_or_load_presentation(
                    presentation_id=presentation_id,
                    file_path=file_path,
                    auto_save=auto_save,
                )
                props = record.presentation.core_properties
                if title:
                    props.title = title
                if author:
                    props.author = author
                if comments:
                    props.comments = comments
                if width_inches > 0 and height_inches > 0:
                    record.presentation.slide_width = Inches(width_inches)
                    record.presentation.slide_height = Inches(height_inches)
                manager.auto_save_if_needed(record)
                return {
                    "status": "success",
                    "action": "set_properties",
                    "presentation_id": record.presentation_id,
                    "title": props.title,
                    "author": props.author,
                }

            else:
                raise ValueError(f"Unknown action '{action}'")

        except Exception as e:
            return {"status": "error", "error": str(e), "action": action}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Inspect Presentation",
            read_only_hint=True,
            destructive_hint=False,
            idempotent_hint=True,
        )
    )
    def inspect_presentation(
        presentation_id: Annotated[
            str | None,
            Field(description="Session ID of open presentation"),
        ] = None,
        file_path: Annotated[
            str | None,
            Field(description="Path to presentation file on disk"),
        ] = None,
        output_format: Annotated[
            Literal["json", "markdown"],
            Field(description="Format of presentation extraction: 'json' or 'markdown'"),
        ] = "json",
    ) -> dict[str, Any] | str:
        """Inspect and extract all content from a presentation in a single call.

        Returns slide count, layout names, speaker notes, all text frames,
        and table rows formatted as structured JSON or readable Markdown.
        """
        try:
            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=False,
            )
            prs = record.presentation

            width_in = round(prs.slide_width.inches, 3)
            height_in = round(prs.slide_height.inches, 3)

            slides_data: list[dict[str, Any]] = []

            for idx, slide in enumerate(prs.slides):
                layout_name = getattr(slide.slide_layout, "name", f"Layout {idx}")
                notes = ""
                if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                    notes = slide.notes_slide.notes_text_frame.text.strip()

                slide_title = None
                elements = []

                for shape_idx, shape in enumerate(slide.shapes):
                    shape_type = shape.shape_type.__str__()
                    l_in = round(shape.left.inches, 2)
                    t_in = round(shape.top.inches, 2)
                    w_in = round(shape.width.inches, 2)
                    h_in = round(shape.height.inches, 2)

                    elem_info: dict[str, Any] = {
                        "shape_index": shape_idx,
                        "name": shape.name,
                        "type": shape_type,
                        "bounds": {"left": l_in, "top": t_in, "width": w_in, "height": h_in},
                    }

                    # Extract text
                    if shape.has_text_frame and shape.text_frame:
                        text_val = shape.text_frame.text.strip()
                        elem_info["text"] = text_val
                        try:
                            is_native_title = shape == slide.shapes.title
                        except Exception:
                            is_native_title = False
                        if is_native_title or (not slide_title and "title" in shape.name.lower()):
                            slide_title = text_val

                    # Extract table
                    if shape.has_table:
                        table = shape.table
                        rows_data: list[list[str]] = []
                        for row in table.rows:
                            rows_data.append([c.text.strip() for c in row.cells])
                        elem_info["table_data"] = rows_data

                    # Extract chart
                    if shape.has_chart:
                        chart = shape.chart
                        chart_dict: dict[str, Any] = {
                            "chart_type": str(chart.chart_type),
                            "has_title": chart.has_title,
                            "title": chart.chart_title.text_frame.text if chart.has_title else None,
                        }
                        elem_info["chart"] = chart_dict

                    elements.append(elem_info)

                if not slide_title:
                    for el in elements:
                        if el.get("text"):
                            first_line = el["text"].split("\n")[0].strip()
                            if first_line and len(first_line) < 120:
                                slide_title = first_line
                                break

                slides_data.append(
                    {
                        "slide_index": idx,
                        "title": slide_title,
                        "layout_name": layout_name,
                        "notes": notes if notes else None,
                        "elements": elements,
                    }
                )

            result = {
                "presentation_id": record.presentation_id,
                "file_path": str(record.file_path) if record.file_path else None,
                "slide_count": len(prs.slides),
                "width_inches": width_in,
                "height_inches": height_in,
                "slides": slides_data,
            }

            if output_format == "markdown":
                md_lines = [
                    f"# Presentation Inspection: {record.presentation_id}",
                    f"- **Total Slides**: {len(prs.slides)}",
                    f'- **Dimensions**: {width_in}" x {height_in}"',
                    f"- **Source File**: {record.file_path or 'In-Memory'}",
                    "",
                ]
                for s in slides_data:
                    md_lines.append(f"## Slide {s['slide_index'] + 1}: {s['title'] or '(Untitled)'}")
                    md_lines.append(f"- *Layout*: {s['layout_name']}")
                    if s["notes"]:
                        md_lines.append(f"- *Speaker Notes*: {s['notes']}")

                    for el in s["elements"]:
                        if "text" in el and el["text"]:
                            md_lines.append(f"  - **{el['name']}**: {el['text']}")
                        elif "table_data" in el:
                            md_lines.append(f"  - **Table ({len(el['table_data'])} rows)**:")
                            for row in el["table_data"]:
                                md_lines.append(f"    | {' | '.join(row)} |")
                        elif "chart" in el:
                            md_lines.append(f"  - **Chart**: {el['chart'].get('title') or el['chart']['chart_type']}")
                    md_lines.append("")
                return "\n".join(md_lines)

            return result

        except Exception as e:
            return {"status": "error", "error": str(e)}
