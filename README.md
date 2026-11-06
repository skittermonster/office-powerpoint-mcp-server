# Office PowerPoint MCP Server (`office-powerpoint-mcp-server`)

[![CI](https://github.com/skittermonster/office-powerpoint-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/skittermonster/office-powerpoint-mcp-server/actions/workflows/ci.yml)
[![Python 3.12+](https://img.shields.io/badge/python-3.12%20%7C%203.13-blue.svg)](https://www.python.org/)
[![FastMCP v4](https://img.shields.io/badge/FastMCP-v4.1.6-purple.svg)](https://github.com/jlowin/fastmcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A production-ready **Model Context Protocol (MCP)** server for automated PowerPoint (`.pptx`) generation, OpenXML-enhanced financial styling, and presentation inspection. Built with Python 3.12+, **FastMCP v4**, **python-pptx**, **lxml**, **Pillow**, and **Pydantic v2**.

Inspired by and functionally referencing `GongRzhe/Office-PowerPoint-MCP-Server`, upgraded to modern 2026 standards with path sandboxing, consolidated tool surface, hybrid state persistence, and native OpenXML table formatting.

---

## Key Architectural Highlights

| Feature | GongRzhe Reference (v2.0) | Office PowerPoint MCP Server (v4) |
|---|---|---|
| **MCP Framework** | FastMCP v1 / MCP Python SDK v1 | **FastMCP v4** (`fastmcp>=4.1.0`) with `ToolAnnotations` hints |
| **Tool Surface** | 32 scattered tools (context bloat) | **14 consolidated, action-dispatched tools** with strict Pydantic schemas |
| **Path Security** | Unrestricted filesystem access | **Strict Path Sandbox (`--allow-dir`)** with traversal (`..`) blocking |
| **State Model** | In-memory only (lost if server restarts) | **Hybrid State**: In-memory sessions **+** direct disk auto-save (`auto_save: true`) |
| **Table Formatting** | Basic python-pptx cells | **Direct OpenXML (`lxml`)**: zebra striping, subtle borders, cell padding |
| **Default Dimensions** | Variable 4:3 / 16:9 | **Widescreen 16:9** (`13.333" x 7.5"`) institutional standard |
| **Package Management** | `pip` / `setup.py` | Modern **`uv`** workflows with Python 3.12+ |

---

## Installation

Install using [`uv`](https://docs.astral.sh/uv/):

```bash
# Clone or navigate to the project directory
cd office-powerpoint-mcp-server

# Create virtual environment and sync dependencies
uv sync
```

---

## Running the MCP Server

The server provides a standard CLI entry point `ppt-mcp-server`:

### 1. Standard I/O (Default for MCP Clients like Claude Desktop, Cursor, Antigravity)
```bash
# Sandboxed to current working directory
uv run ppt-mcp-server

# Sandboxed to specific directories
uv run ppt-mcp-server stdio --allow-dir /Users/username/decks --allow-dir /Volumes/Presentations
```

### 2. HTTP Transport (Streamable HTTP / SSE)
```bash
uv run ppt-mcp-server http --host 127.0.0.1 --port 8000 --allow-dir /Users/username/decks
```

---

## Client Configuration

### Claude Desktop / Antigravity (`claude_desktop_config.json` / `mcp-config.json`)
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/office-powerpoint-mcp-server",
        "ppt-mcp-server",
        "stdio",
        "--allow-dir",
        "/path/to/allowed/decks"
      ]
    }
  }
}
```

---

## Tool Reference (14 Consolidated Tools)

All tools expose MCP `ToolAnnotations` (`read_only_hint`, `destructive_hint`, `idempotent_hint`).

### 1. Lifecycle & Ingestion (`tools/lifecycle.py`)
- **`manage_presentation`**: Actions: `create`, `open`, `save`, `close`, `set_properties`.
  - Sets widescreen 16:9 (`13.333" x 7.5"`) by default.
  - Supports hybrid in-memory IDs and direct disk paths with `auto_save: true`.
- **`inspect_presentation`** *(Read-Only)*: Ingests an entire presentation in one call.
  - Returns slide counts, layout names, speaker notes, all text frames, and table rows as structured JSON or publication Markdown.

### 2. Themes & Institutional Styling (`tools/design.py` & `templates.json`)
- **`apply_theme`**: Applies institutional color schemes, slide background, and font hierarchies across slides or tables.
  - Supported Themes: `institutional_navy`, `modern_blue`, `corporate_gray`, `elegant_green`, `warm_red`.
- **`format_master_footer`**: Adds institutional divider line, confidentiality notice, current date, and dynamic page numbers (`Slide X of Y`) across slides.

### 3. Template-Driven Slide Generation (`tools/slides.py`)
- **`add_slide_from_template`**: Builds pre-calculated coordinate slides from `templates.json`:
  1. `title_slide`: Authoritative title slide with decorative accent bar, subtitle, author, and date.
  2. `executive_summary`: Dual strategic cards for observations and recommendations.
  3. `two_column_comparison`: Side-by-side comparison cards (Option A vs Option B / Current vs Target).
  4. `kpi_metrics_grid`: 4 high-impact metric cards with prominent numbers, labels, and YoY trends.
  5. `table_with_commentary`: Left-side financial data table with right-side commentary card.
  6. `chart_with_takeaways`: Left-side native chart with right-side analyst takeaways card.
- **`manage_slides`**: Slide operations: `add`, `duplicate`, `delete`, `reorder`, `populate_placeholder`.
- **`list_templates`** *(Read-Only)*: Lists all available layout templates and color schemes.

### 4. Financial & Data Tables (`tools/tables.py`)
- **`add_data_table`**: Creates publication-ready financial tables:
  - Header row: Theme primary background (`#1B365D`) with bold white text.
  - Alternating rows: Slate gray zebra striping (`#F1F5F9`).
  - Subtle OpenXML horizontal borders (`#CBD5E1`) via `lxml`.
  - Left-aligned text in Column 0, right-aligned numbers in Columns $\ge 1$.
  - Optional bottom source / footnote line.
- **`update_table_cells`**: Fine-grained cell updates (text, colors, bold, alignment, font sizes).

### 5. Text & Shape Management (`tools/content.py`)
- **`manage_text`**: Add, update, or clear text boxes and multi-level bulleted lists (levels 0–4) with run-level overrides and auto-fit shrink-to-fit sizing.
- **`add_bullet_slide`**: Quick creation of institutional structured bullet slide with header and divider.
- **`manage_shapes`**: Create cards, callout containers, arrows, rectangles, and rounded containers.

### 6. Media & Charts (`tools/media_charts.py`)
- **`add_native_chart`**: Creates native editable PowerPoint charts (`COLUMN_CLUSTERED`, `LINE`, `PIE`, `BAR_CLUSTERED`) using `CategoryChartData`.
- **`add_image`**: Inserts images from a validated local file path or base64 string (e.g. from `matplotlib`), locking aspect ratio within a bounding box `(left, top, max_width, max_height)`.

---

## Institutional Navy Palette Specifications

```yaml
Primary Accent / Header: '#1B365D' (Muted Institutional Navy)
Secondary Accent:        '#0F172A' (Slate Navy)
Zebra Row / Card Fill:   '#F1F5F9' (Slate Gray)
Border Lines:            '#CBD5E1' (Light Slate Border)
Body Text:               '#1A1A1A' (Dark Charcoal)
Muted Subtext:           '#64748B' (Muted Slate)
Background:              '#FFFFFF' (Pure White)
Fonts:
  Title: 'Helvetica'
  Body:  'Arial'
```

---

## Quality & Automated Verification

Run tests and code formatting checks:

```bash
# Run full pytest test suite
uv run pytest -v

# Run Ruff linter and formatter checks
uv run ruff check .
uv run ruff format --check .
```

---

## License

MIT License.
