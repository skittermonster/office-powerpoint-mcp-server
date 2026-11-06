"""Media and native chart creation tools for PowerPoint MCP Server."""

import base64
import io
from typing import Annotated, Any, Literal

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from PIL import Image
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches
from pydantic import Field

from powerpoint_mcp.models import ChartSeriesInput, SupportedChartType
from powerpoint_mcp.state import SessionManager, session_manager
from powerpoint_mcp.utils.geometry import fit_image_in_bounding_box


def register_media_chart_tools(mcp: FastMCP, manager: SessionManager = session_manager) -> None:
    """Register media and native chart creation tools with FastMCP."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Native Chart",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def add_native_chart(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide to place chart on"),
        ],
        chart_type: Annotated[
            SupportedChartType,
            Field(description="Chart type: COLUMN_CLUSTERED, LINE, PIE, or BAR_CLUSTERED"),
        ],
        categories: Annotated[
            list[str],
            Field(description="List of category names along X-axis / slices"),
        ],
        series: Annotated[
            list[ChartSeriesInput],
            Field(description="List of data series with series name and numerical values"),
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
            Field(description="Chart width in inches"),
        ] = 8.0,
        height: Annotated[
            float,
            Field(description="Chart height in inches"),
        ] = 4.8,
        title: Annotated[
            str | None,
            Field(description="Optional chart header title"),
        ] = None,
        has_legend: Annotated[
            bool,
            Field(description="Whether to render chart legend"),
        ] = True,
        legend_position: Annotated[
            Literal["top", "bottom", "left", "right"],
            Field(description="Legend position: top, bottom, left, right"),
        ] = "right",
        show_data_labels: Annotated[
            bool,
            Field(description="Whether to display numeric data labels on chart points/bars"),
        ] = False,
        show_gridlines: Annotated[
            bool,
            Field(description="Whether to show major gridlines on value axis"),
        ] = True,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Create a native, editable PowerPoint chart.

        Supports COLUMN_CLUSTERED, LINE, PIE, and BAR_CLUSTERED charts with
        custom series data, legend positioning, gridline toggling, and data labels.
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

            type_map = {
                "COLUMN_CLUSTERED": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "LINE": XL_CHART_TYPE.LINE,
                "PIE": XL_CHART_TYPE.PIE,
                "BAR_CLUSTERED": XL_CHART_TYPE.BAR_CLUSTERED,
            }
            xl_type = type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

            chart_data = CategoryChartData()
            chart_data.categories = categories

            for s in series:
                chart_data.add_series(s.name, tuple(s.values))

            chart_shape = slide.shapes.add_chart(
                xl_type,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height),
                chart_data,
            )
            chart = chart_shape.chart

            # Configure title
            if title:
                chart.has_title = True
                chart.chart_title.text_frame.text = title
            else:
                chart.has_title = False

            # Configure legend
            chart.has_legend = has_legend
            if has_legend:
                pos_map = {
                    "top": XL_LEGEND_POSITION.TOP,
                    "bottom": XL_LEGEND_POSITION.BOTTOM,
                    "left": XL_LEGEND_POSITION.LEFT,
                    "right": XL_LEGEND_POSITION.RIGHT,
                }
                chart.legend.position = pos_map.get(legend_position, XL_LEGEND_POSITION.RIGHT)

            # Configure data labels
            if show_data_labels and len(chart.plots) > 0:
                chart.plots[0].has_data_labels = True

            # Configure gridlines if chart has a value axis (e.g. not PIE)
            if chart_type != "PIE":
                try:
                    chart.value_axis.has_major_gridlines = show_gridlines
                except Exception:
                    pass

            manager.auto_save_if_needed(record)

            shape_index = len(slide.shapes) - 1
            return {
                "status": "success",
                "slide_index": slide_index,
                "shape_index": shape_index,
                "chart_type": chart_type,
                "series_count": len(series),
                "categories_count": len(categories),
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Add Image",
            read_only_hint=False,
            destructive_hint=False,
            idempotent_hint=False,
        )
    )
    def add_image(
        slide_index: Annotated[
            int,
            Field(ge=0, description="0-based index of slide to place image on"),
        ],
        image_path: Annotated[
            str | None,
            Field(description="Path to image file within allowed directory sandbox"),
        ] = None,
        image_base64: Annotated[
            str | None,
            Field(description="Base64 encoded image string (e.g. from matplotlib or camera)"),
        ] = None,
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
            Field(description="Left coordinate of bounding box in inches"),
        ] = 1.0,
        top: Annotated[
            float,
            Field(description="Top coordinate of bounding box in inches"),
        ] = 1.8,
        max_width: Annotated[
            float,
            Field(description="Maximum allowable width in inches"),
        ] = 6.0,
        max_height: Annotated[
            float,
            Field(description="Maximum allowable height in inches"),
        ] = 4.5,
        preserve_aspect_ratio: Annotated[
            bool,
            Field(description="Whether to lock aspect ratio and fit cleanly in bounding box"),
        ] = True,
        auto_save: Annotated[
            bool,
            Field(description="Whether to automatically save changes to file_path on disk"),
        ] = True,
    ) -> dict[str, Any]:
        """Insert an image into a slide from a local file path or base64 string.

        Automatically preserves aspect ratio within the specified bounding box
        (left, top, max_width, max_height).
        """
        try:
            if not image_path and not image_base64:
                raise ValueError("Either image_path or image_base64 must be provided")

            record = manager.get_or_load_presentation(
                presentation_id=presentation_id,
                file_path=file_path,
                auto_save=auto_save,
            )
            prs = record.presentation

            if slide_index < 0 or slide_index >= len(prs.slides):
                raise IndexError(f"slide_index {slide_index} out of range")

            slide = prs.slides[slide_index]

            if image_base64:
                # Clean base64 string if data URL prefix exists
                raw_b64 = image_base64
                if "," in raw_b64:
                    raw_b64 = raw_b64.split(",", 1)[1]
                raw_bytes = base64.b64decode(raw_b64.strip())
                img_io = io.BytesIO(raw_bytes)
                pil_img = Image.open(img_io)
                img_w, img_h = pil_img.size
                img_io.seek(0)
                image_source = img_io
            else:
                resolved_img = manager.validate_path(image_path, allow_nonexistent=False, must_be_file=True)
                with Image.open(str(resolved_img)) as pil_img:
                    img_w, img_h = pil_img.size
                image_source = str(resolved_img)

            p_left, p_top, p_width, p_height = fit_image_in_bounding_box(
                img_w,
                img_h,
                left=left,
                top=top,
                max_width=max_width,
                max_height=max_height,
                preserve_aspect_ratio=preserve_aspect_ratio,
            )

            slide.shapes.add_picture(
                image_source,
                Inches(p_left),
                Inches(p_top),
                Inches(p_width),
                Inches(p_height),
            )

            manager.auto_save_if_needed(record)

            shape_index = len(slide.shapes) - 1
            return {
                "status": "success",
                "slide_index": slide_index,
                "shape_index": shape_index,
                "placed_bounds": {
                    "left_inches": p_left,
                    "top_inches": p_top,
                    "width_inches": p_width,
                    "height_inches": p_height,
                },
                "original_dimensions": {"width_px": img_w, "height_px": img_h},
            }

        except Exception as e:
            return {"status": "error", "error": str(e)}
