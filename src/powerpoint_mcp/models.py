"""Pydantic input and output models for Office PowerPoint MCP Server."""

from typing import Any, Literal

from pydantic import BaseModel, Field

# Supported themes
ThemeName = Literal[
    "institutional_navy",
    "modern_blue",
    "corporate_gray",
    "elegant_green",
    "warm_red",
]

# Supported templates
TemplateName = Literal[
    "title_slide",
    "executive_summary",
    "two_column_comparison",
    "kpi_metrics_grid",
    "table_with_commentary",
    "chart_with_takeaways",
]

# Chart types
SupportedChartType = Literal[
    "COLUMN_CLUSTERED",
    "LINE",
    "PIE",
    "BAR_CLUSTERED",
]

# Text alignments
TextAlignment = Literal["left", "center", "right", "justify"]

# Shape types
SupportedShapeType = Literal[
    "rectangle",
    "rounded_rectangle",
    "card",
    "callout",
    "arrow",
    "oval",
]


class ThemeFonts(BaseModel):
    title: str = "Helvetica"
    body: str = "Arial"


class ThemePalette(BaseModel):
    name: str
    primary: str
    secondary: str
    accent1: str
    accent2: str
    light: str
    zebra: str
    text: str
    muted_text: str
    background: str
    card_bg: str = "#FFFFFF"
    border: str
    fonts: ThemeFonts = Field(default_factory=ThemeFonts)


class BulletItem(BaseModel):
    text: str = Field(description="Bullet text content")
    level: int = Field(default=0, ge=0, le=4, description="Bullet indentation level (0 to 4)")
    bold: bool = Field(default=False, description="Whether bullet is bold")
    italic: bool = Field(default=False, description="Whether bullet is italic")
    color: str | None = Field(default=None, description="Hex color override e.g. '#1B365D'")
    size_pt: float | None = Field(default=None, description="Font size in points")


class TableCellUpdate(BaseModel):
    row: int = Field(ge=0, description="0-based row index")
    col: int = Field(ge=0, description="0-based column index")
    text: str | None = Field(default=None, description="New cell text")
    font_size: float | None = Field(default=None, description="Font size in pt")
    bold: bool | None = Field(default=None, description="Bold styling")
    italic: bool | None = Field(default=None, description="Italic styling")
    text_color: str | None = Field(default=None, description="Hex text color e.g. '#1A1A1A'")
    bg_color: str | None = Field(default=None, description="Hex cell background color e.g. '#F1F5F9'")
    alignment: TextAlignment | None = Field(default=None, description="Text alignment")


class ChartSeriesInput(BaseModel):
    name: str = Field(description="Name of the data series")
    values: list[float | int] = Field(description="Numeric data values corresponding to categories")


class KpiMetricItem(BaseModel):
    value: str = Field(description="Main prominent metric string, e.g. '$14.2M' or '+85%'")
    label: str = Field(description="KPI title / label, e.g. 'Annual Recurring Revenue'")
    subtext: str | None = Field(default=None, description="Subtext, trend, or comparison, e.g. '+24% YoY'")
    badge: str | None = Field(default=None, description="Optional small badge or category")


class SlideElementInspection(BaseModel):
    shape_index: int
    shape_type: str
    name: str
    left_inches: float
    top_inches: float
    width_inches: float
    height_inches: float
    text: str | None = None
    table_data: list[list[str]] | None = None
    chart_info: dict[str, Any] | None = None


class SlideInspection(BaseModel):
    slide_index: int
    layout_name: str
    title: str | None = None
    speaker_notes: str | None = None
    elements: list[SlideElementInspection] = Field(default_factory=list)


class PresentationInspection(BaseModel):
    presentation_id: str
    file_path: str | None = None
    slide_count: int
    width_inches: float
    height_inches: float
    slides: list[SlideInspection]
