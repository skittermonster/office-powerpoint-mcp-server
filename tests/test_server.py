"""Comprehensive test suite for Office PowerPoint MCP Server."""

import base64
import io
from pathlib import Path

import pytest
from fastmcp import Client
from PIL import Image

from powerpoint_mcp.server import create_server
from powerpoint_mcp.state import SessionManager


@pytest.fixture
def temp_sandbox(tmp_path: Path):
    """Fixture providing an isolated sandboxed directory and session manager."""
    allowed = tmp_path / "allowed_decks"
    allowed.mkdir(parents=True, exist_ok=True)
    manager = SessionManager(allowed_dirs=[allowed])
    test_mcp = create_server(manager)
    return allowed, manager, test_mcp


@pytest.mark.asyncio
async def test_presentation_lifecycle(temp_sandbox):
    """Test creating, configuring, saving, closing, and reopening presentations."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "lifecycle_test.pptx"

    async with Client(server) as client:
        # 1. Create presentation
        create_res = await client.call_tool(
            "manage_presentation",
            {
                "action": "create",
                "file_path": str(deck_path),
                "auto_save": False,
                "title": "Strategic Review 2026",
                "author": "Investment Committee",
            },
        )
        data = create_res.data
        assert data["status"] == "success"
        pid = data["presentation_id"]
        assert data["dimensions"]["width_inches"] == 13.333
        assert data["dimensions"]["height_inches"] == 7.5

        # 2. Add a slide
        add_slide_res = await client.call_tool(
            "manage_slides",
            {"action": "add", "presentation_id": pid, "layout_index": 6},
        )
        assert add_slide_res.data["status"] == "success"
        assert add_slide_res.data["slide_count"] == 1

        # 3. Save to disk
        save_res = await client.call_tool(
            "manage_presentation",
            {"action": "save", "presentation_id": pid, "file_path": str(deck_path)},
        )
        assert save_res.data["status"] == "success"
        assert deck_path.exists()

        # 4. Close presentation
        close_res = await client.call_tool(
            "manage_presentation",
            {"action": "close", "presentation_id": pid},
        )
        assert close_res.data["status"] == "success"

        # 5. Open from disk
        open_res = await client.call_tool(
            "manage_presentation",
            {"action": "open", "file_path": str(deck_path)},
        )
        assert open_res.data["status"] == "success"
        assert open_res.data["slide_count"] == 1


@pytest.mark.asyncio
async def test_path_sandbox_security(temp_sandbox):
    """Verify that path traversal and unauthorized directory access are blocked."""
    allowed_dir, manager, server = temp_sandbox

    # Test direct manager validate_path
    with pytest.raises(ValueError, match="Access denied"):
        manager.validate_path("/etc/passwd")

    with pytest.raises(ValueError, match="Access denied"):
        manager.validate_path(allowed_dir / ".." / "unauthorized.pptx")

    # Test through MCP tool
    async with Client(server) as client:
        res = await client.call_tool(
            "manage_presentation",
            {"action": "create", "file_path": "/tmp/forbidden_escape.pptx"},
        )
        assert res.data["status"] == "error"
        assert "Access denied" in res.data["error"]


@pytest.mark.asyncio
async def test_template_slides_generation(temp_sandbox):
    """Test generating all 6 pre-calculated coordinate templates."""
    allowed_dir, manager, server = temp_sandbox

    async with Client(server) as client:
        # Create deck
        create_res = await client.call_tool(
            "manage_presentation",
            {"action": "create", "presentation_id": "templates_deck"},
        )
        pid = create_res.data["presentation_id"]

        # 1. list_templates
        templates_list = await client.call_tool("list_templates", {})
        assert templates_list.data["status"] == "success"
        assert "title_slide" in templates_list.data["templates"]
        assert "institutional_navy" in templates_list.data["color_schemes"]

        # 2. Add title_slide
        t1 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "title_slide",
                "title": "Global Logistics Platform",
                "subtitle": "Institutional Investment Memorandum",
                "elements_data": {
                    "author": "Private Equity Growth Fund",
                    "date": "November 2026",
                },
                "color_scheme": "institutional_navy",
            },
        )
        assert t1.data["status"] == "success"
        assert t1.data["slide_index"] == 0

        # 3. Add executive_summary
        t2 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "executive_summary",
                "title": "Executive Summary & Investment Highlights",
                "subtitle": "High-conviction market leadership opportunity",
                "elements_data": {
                    "left_title": "Core Strengths",
                    "left_bullets": [
                        "Category leadership with 42% market share.",
                        "Defensible proprietary software moat.",
                    ],
                    "right_title": "Value Creation Plan",
                    "right_bullets": [
                        "Expand cross-border European coverage.",
                        "Drive EBITDA margin from 22% to 30%.",
                    ],
                },
            },
        )
        assert t2.data["status"] == "success"
        assert t2.data["slide_index"] == 1

        # 4. Add two_column_comparison
        t3 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "two_column_comparison",
                "title": "Architectural Evolution",
                "elements_data": {
                    "col1_title": "Legacy On-Premise",
                    "col1_bullets": ["Monolithic stack", "Manual deployment"],
                    "col2_title": "Modern Cloud-Native",
                    "col2_bullets": ["Event-driven microservices", "Automated CI/CD"],
                },
            },
        )
        assert t3.data["status"] == "success"
        assert t3.data["slide_index"] == 2

        # 5. Add kpi_metrics_grid
        t4 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "kpi_metrics_grid",
                "title": "Operational KPI Dashboard",
                "elements_data": {
                    "metrics": [
                        {"value": "$124M", "label": "ARR", "subtext": "+34% YoY"},
                        {"value": "78.2%", "label": "Gross Margin", "subtext": "+420 bps"},
                        {"value": "131%", "label": "Net Retention", "subtext": "Best in Class"},
                        {"value": "3.8x", "label": "LTV / CAC", "subtext": "Payback 9 mos"},
                    ]
                },
            },
        )
        assert t4.data["status"] == "success"
        assert t4.data["slide_index"] == 3

        # 6. Add table_with_commentary
        t5 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "table_with_commentary",
                "title": "Unit Economics by Segment",
                "elements_data": {
                    "table_headers": ["Segment", "Revenue", "Contribution"],
                    "table_rows": [
                        ["Enterprise", "$82M", "68%"],
                        ["Mid-Market", "$42M", "54%"],
                    ],
                    "commentary_title": "Analyst Perspective",
                    "commentary_bullets": ["Enterprise segment accounts for 75% of new bookings."],
                },
            },
        )
        assert t5.data["status"] == "success"
        assert t5.data["slide_index"] == 4

        # 7. Add chart_with_takeaways
        t6 = await client.call_tool(
            "add_slide_from_template",
            {
                "presentation_id": pid,
                "template_name": "chart_with_takeaways",
                "title": "Quarterly Performance Trajectory",
                "elements_data": {
                    "chart_type": "COLUMN_CLUSTERED",
                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                    "series": [{"name": "2026 Actual", "values": [22, 28, 35, 42]}],
                    "takeaways_title": "Key Drivers",
                    "takeaways_bullets": ["Record expansion in fourth quarter."],
                },
            },
        )
        assert t6.data["status"] == "success"
        assert t6.data["slide_index"] == 5


@pytest.mark.asyncio
async def test_financial_table_and_openxml(temp_sandbox):
    """Test institutional data table creation, auto column widths, zebra shading, and cell updates."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "table_test.pptx"

    async with Client(server) as client:
        # Create presentation
        await client.call_tool(
            "manage_presentation",
            {"action": "create", "file_path": str(deck_path), "auto_save": True},
        )

        # Add slide
        await client.call_tool("manage_slides", {"action": "add"})

        # Add formatted table
        headers = ["Fiscal Year", "Revenue ($M)", "Gross Profit ($M)", "EBITDA ($M)", "Margin"]
        rows = [
            ["FY2023", "45.0", "31.5", "8.2", "18.2%"],
            ["FY2024", "68.2", "49.1", "15.4", "22.6%"],
            ["FY2025", "102.5", "76.9", "28.7", "28.0%"],
            ["FY2026E", "150.0", "115.5", "48.0", "32.0%"],
        ]
        tbl_res = await client.call_tool(
            "add_data_table",
            {
                "slide_index": 0,
                "headers": headers,
                "rows": rows,
                "theme": "institutional_navy",
                "footnote": "SEC Filings & Management Financial Models",
                "highlight_first_column": True,
            },
        )
        assert tbl_res.data["status"] == "success"
        assert tbl_res.data["rows"] == 5
        assert tbl_res.data["columns"] == 5
        shape_idx = tbl_res.data["shape_index"]

        # Update specific cells
        update_res = await client.call_tool(
            "update_table_cells",
            {
                "slide_index": 0,
                "shape_index": shape_idx,
                "updates": [
                    {
                        "row": 4,
                        "col": 4,
                        "text": "34.5%",
                        "bold": True,
                        "text_color": "#2563EB",
                    }
                ],
            },
        )
        assert update_res.data["status"] == "success"
        assert update_res.data["cells_updated"] == 1


@pytest.mark.asyncio
async def test_native_charts_and_images(temp_sandbox):
    """Test native chart types and image insertion via path and base64 with aspect ratio locking."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "media_test.pptx"

    # Create dummy PNG image on disk
    img_file = allowed_dir / "test_logo.png"
    img = Image.new("RGB", (400, 200), color=(27, 54, 93))
    img.save(img_file, format="PNG")

    # Create base64 image
    b64_buffer = io.BytesIO()
    Image.new("RGB", (300, 300), color=(37, 99, 235)).save(b64_buffer, format="PNG")
    b64_str = base64.b64encode(b64_buffer.getvalue()).decode()

    async with Client(server) as client:
        await client.call_tool(
            "manage_presentation",
            {"action": "create", "file_path": str(deck_path), "auto_save": True},
        )

        # Slide 0: Column chart
        await client.call_tool("manage_slides", {"action": "add"})
        c1 = await client.call_tool(
            "add_native_chart",
            {
                "slide_index": 0,
                "chart_type": "COLUMN_CLUSTERED",
                "categories": ["North America", "EMEA", "APAC"],
                "series": [
                    {"name": "2025", "values": [40, 25, 15]},
                    {"name": "2026", "values": [65, 42, 30]},
                ],
                "title": "Regional Revenue Growth",
                "has_legend": True,
                "legend_position": "right",
            },
        )
        assert c1.data["status"] == "success"

        # Slide 1: Pie chart
        await client.call_tool("manage_slides", {"action": "add"})
        c2 = await client.call_tool(
            "add_native_chart",
            {
                "slide_index": 1,
                "chart_type": "PIE",
                "categories": ["Enterprise", "Mid-Market", "SMB"],
                "series": [{"name": "Market Share", "values": [55, 30, 15]}],
                "title": "Customer Segmentation",
            },
        )
        assert c2.data["status"] == "success"

        # Slide 2: Image from file
        await client.call_tool("manage_slides", {"action": "add"})
        img_res1 = await client.call_tool(
            "add_image",
            {
                "slide_index": 2,
                "image_path": str(img_file),
                "max_width": 4.0,
                "max_height": 2.0,
                "preserve_aspect_ratio": True,
            },
        )
        assert img_res1.data["status"] == "success"
        bounds = img_res1.data["placed_bounds"]
        assert bounds["width_inches"] == 4.0
        assert bounds["height_inches"] == 2.0

        # Slide 2: Image from base64
        img_res2 = await client.call_tool(
            "add_image",
            {
                "slide_index": 2,
                "image_base64": b64_str,
                "left": 6.0,
                "top": 1.8,
                "max_width": 3.0,
                "max_height": 3.0,
                "preserve_aspect_ratio": True,
            },
        )
        assert img_res2.data["status"] == "success"


@pytest.mark.asyncio
async def test_text_and_shape_management(temp_sandbox):
    """Test text boxes, auto-fit, multi-level bullets, shapes, and bullet slides."""
    allowed_dir, manager, server = temp_sandbox

    async with Client(server) as client:
        await client.call_tool("manage_presentation", {"action": "create"})
        await client.call_tool("manage_slides", {"action": "add"})

        # 1. Add text box
        txt_res = await client.call_tool(
            "manage_text",
            {
                "slide_index": 0,
                "action": "add",
                "text": "Strategic Outlook",
                "left": 1.0,
                "top": 1.0,
                "width": 6.0,
                "height": 1.0,
                "font_size": 24,
                "bold": True,
            },
        )
        assert txt_res.data["status"] == "success"
        shape_idx = txt_res.data["shape_index"]

        # 2. Update with bullets
        update_txt = await client.call_tool(
            "manage_text",
            {
                "slide_index": 0,
                "action": "update",
                "shape_index": shape_idx,
                "bullet_points": [
                    {"text": "First strategic priority", "level": 0, "bold": True},
                    {"text": "Sub-point execution detail", "level": 1},
                ],
            },
        )
        assert update_txt.data["status"] == "success"

        # 3. Add visual card shape
        shape_res = await client.call_tool(
            "manage_shapes",
            {
                "slide_index": 0,
                "action": "add",
                "shape_type": "card",
                "left": 1.0,
                "top": 3.0,
                "width": 4.0,
                "height": 2.5,
                "fill_color": "#F1F5F9",
                "border_color": "#CBD5E1",
                "text": "Card Container",
                "bold": True,
            },
        )
        assert shape_res.data["status"] == "success"
        new_shape_idx = shape_res.data["shape_index"]

        # 4. Delete shape
        del_shape = await client.call_tool(
            "manage_shapes",
            {
                "slide_index": 0,
                "action": "delete",
                "shape_index": new_shape_idx,
            },
        )
        assert del_shape.data["status"] == "success"

        # 5. Add quick bullet slide
        b_slide = await client.call_tool(
            "add_bullet_slide",
            {
                "title": "Key Operational Priorities",
                "subtitle": "Detailed Q1 roadmap",
                "bullet_points": [
                    "Accelerate customer acquisition velocity",
                    "Strengthen compliance framework",
                ],
                "theme": "institutional_navy",
            },
        )
        assert b_slide.data["status"] == "success"
        assert b_slide.data["bullet_count"] == 2


@pytest.mark.asyncio
async def test_full_roundtrip_inspection(temp_sandbox):
    """End-to-end verification: build presentation, save to disk, and inspect in JSON & Markdown."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "institutional_pitchbook.pptx"

    async with Client(server) as client:
        # Create presentation
        await client.call_tool(
            "manage_presentation",
            {
                "action": "create",
                "file_path": str(deck_path),
                "auto_save": True,
                "title": "Alpha Growth Partners Pitchbook",
            },
        )

        # Slide 0: Title Slide
        await client.call_tool(
            "add_slide_from_template",
            {
                "template_name": "title_slide",
                "title": "Alpha Growth Partners",
                "subtitle": "Fund IV Investor Presentation",
                "elements_data": {"author": "General Partner Team", "date": "Q4 2026"},
            },
        )

        # Slide 1: Executive Summary
        await client.call_tool(
            "add_slide_from_template",
            {
                "template_name": "executive_summary",
                "title": "Executive Summary",
                "elements_data": {
                    "left_title": "Market Opportunity",
                    "left_bullets": ["$50B addressable software market expanding at 22% CAGR."],
                    "right_title": "Investment Thesis",
                    "right_bullets": ["Disciplined buyout strategy targeting industry leaders."],
                },
            },
        )

        # Slide 2: Formatted Financial Table
        await client.call_tool("manage_slides", {"action": "add"})
        await client.call_tool(
            "add_data_table",
            {
                "slide_index": 2,
                "headers": ["Metric", "2024", "2025", "2026E"],
                "rows": [
                    ["Revenue ($M)", "$120.0", "$165.0", "$220.0"],
                    ["EBITDA ($M)", "$28.0", "$42.0", "$65.0"],
                ],
                "footnote": "Audited Financials",
            },
        )

        # Slide 3: Native Column Chart
        await client.call_tool("manage_slides", {"action": "add"})
        await client.call_tool(
            "add_native_chart",
            {
                "slide_index": 3,
                "chart_type": "COLUMN_CLUSTERED",
                "categories": ["2024", "2025", "2026E"],
                "series": [{"name": "Fund Distributions", "values": [15, 35, 70]}],
                "title": "Cumulative Capital Distributions",
            },
        )

        # Apply institutional styling & master footer
        await client.call_tool("apply_theme", {"theme_name": "institutional_navy"})
        await client.call_tool(
            "format_master_footer",
            {
                "footer_text": "Confidential & Proprietary",
                "include_slide_numbers": True,
                "include_date": True,
            },
        )

        # Verify inspection as JSON
        inspect_json = await client.call_tool(
            "inspect_presentation",
            {"file_path": str(deck_path), "output_format": "json"},
        )
        data = inspect_json.data
        assert data["slide_count"] == 4
        assert data["slides"][0]["title"] == "Alpha Growth Partners"

        # Verify table data roundtrip
        slide_2_elements = data["slides"][2]["elements"]
        table_found = False
        for el in slide_2_elements:
            if "table_data" in el:
                table_found = True
                assert el["table_data"][0] == ["Metric", "2024", "2025", "2026E"]
                assert el["table_data"][1] == ["Revenue ($M)", "$120.0", "$165.0", "$220.0"]
        assert table_found

        # Verify chart data roundtrip
        slide_3_elements = data["slides"][3]["elements"]
        chart_found = any("chart" in el for el in slide_3_elements)
        assert chart_found

        # Verify inspection as Markdown
        inspect_md = await client.call_tool(
            "inspect_presentation",
            {"file_path": str(deck_path), "output_format": "markdown"},
        )
        md_text = inspect_md.data
        assert isinstance(md_text, str)
        assert "# Presentation Inspection" in md_text
        assert "Alpha Growth Partners" in md_text
        assert "Audited Financials" in md_text or "Metric" in md_text


@pytest.mark.asyncio
async def test_slide_manipulation_actions(temp_sandbox):
    """Test duplicating, reordering, deleting, and placeholder population."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "slide_actions.pptx"

    async with Client(server) as client:
        await client.call_tool(
            "manage_presentation",
            {"action": "create", "file_path": str(deck_path), "auto_save": True},
        )
        # Add 3 slides
        await client.call_tool("manage_slides", {"action": "add"})
        await client.call_tool("manage_slides", {"action": "add"})
        await client.call_tool("manage_slides", {"action": "add"})

        # Duplicate slide 1
        dup = await client.call_tool("manage_slides", {"action": "duplicate", "slide_index": 1})
        assert dup.data["status"] == "success"
        assert dup.data["slide_count"] == 4

        # Reorder slide 3 to index 0
        reorder = await client.call_tool("manage_slides", {"action": "reorder", "slide_index": 3, "to_index": 0})
        assert reorder.data["status"] == "success"

        # Delete slide 0
        deleted = await client.call_tool("manage_slides", {"action": "delete", "slide_index": 0})
        assert deleted.data["status"] == "success"
        assert deleted.data["slide_count"] == 3


@pytest.mark.asyncio
async def test_hybrid_state_direct_autosave(temp_sandbox):
    """Verify that auto_save=True modifies the file on disk between independent calls."""
    allowed_dir, manager, server = temp_sandbox
    deck_path = allowed_dir / "autosave_deck.pptx"

    async with Client(server) as client:
        # Create on disk
        await client.call_tool(
            "manage_presentation",
            {"action": "create", "file_path": str(deck_path), "auto_save": True},
        )
        assert deck_path.exists()
        initial_mtime = deck_path.stat().st_mtime_ns

        # Add slide directly to file_path with auto_save
        res = await client.call_tool(
            "add_bullet_slide",
            {
                "file_path": str(deck_path),
                "auto_save": True,
                "title": "Auto Saved Headline",
                "bullet_points": ["Persisted directly on disk without explicit save call"],
            },
        )
        assert res.data["status"] == "success"
        new_mtime = deck_path.stat().st_mtime_ns
        assert new_mtime >= initial_mtime

        # Inspect disk file directly
        inspect_res = await client.call_tool(
            "inspect_presentation",
            {"file_path": str(deck_path), "output_format": "json"},
        )
        assert inspect_res.data["slide_count"] == 1
        assert "Auto Saved Headline" in str(inspect_res.data)
