"""OpenXML manipulation utilities using lxml for PowerPoint elements.

Provides fine-grained styling that python-pptx does not natively support:
- Individual table cell borders (top, bottom, left, right with specific colors/widths)
- Table cell margins / padding
- Table cell background shading
- Slide background coloring via OpenXML p:bg elements
"""

from typing import Literal

from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.slide import Slide
from pptx.table import _Cell


def normalize_hex_color(hex_str: str) -> str:
    """Normalize hex color string to 6-character uppercase without #."""
    cleaned = hex_str.strip().lstrip("#")
    if len(cleaned) == 3:
        cleaned = "".join([c * 2 for c in cleaned])
    if len(cleaned) != 6:
        return "000000"
    return cleaned.upper()


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert hex color string to python-pptx RGBColor."""
    cleaned = normalize_hex_color(hex_str)
    r = int(cleaned[0:2], 16)
    g = int(cleaned[2:4], 16)
    b = int(cleaned[4:6], 16)
    return RGBColor(r, g, b)


def set_cell_border(
    cell: _Cell,
    side: Literal["top", "bottom", "left", "right"],
    color_hex: str = "CBD5E1",
    width_pt: float = 1.0,
    style: Literal["solid", "none"] = "solid",
) -> None:
    """Set border styling on an individual table cell edge.

    Args:
        cell: The python-pptx _Cell object
        side: Which border side to configure ('top', 'bottom', 'left', 'right')
        color_hex: Hex color for the border
        width_pt: Border width in points (1 pt = 12700 EMUs)
        style: Border style ('solid' or 'none')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    side_map = {
        "top": ("a:lnT", "lnT"),
        "bottom": ("a:lnB", "lnB"),
        "left": ("a:lnL", "lnL"),
        "right": ("a:lnR", "lnR"),
    }
    tag_prefix, xml_tag = side_map[side]

    # Remove any existing border on this side
    for old_elem in tcPr.findall(qn(tag_prefix)):
        tcPr.remove(old_elem)

    if style == "none" or width_pt <= 0:
        border_xml = f'<{tag_prefix} {nsdecls("a")} w="0"><a:noFill/></{tag_prefix}>'
    else:
        width_emu = int(width_pt * 12700)
        clean_hex = normalize_hex_color(color_hex)
        border_xml = (
            f'<{tag_prefix} {nsdecls("a")} w="{width_emu}" cmpd="s">'
            f'  <a:solidFill><a:srgbClr val="{clean_hex}"/></a:solidFill>'
            f"</{tag_prefix}>"
        )

    border_element = parse_xml(border_xml)
    tcPr.append(border_element)


def set_cell_borders(
    cell: _Cell,
    top: dict | None = None,
    bottom: dict | None = None,
    left: dict | None = None,
    right: dict | None = None,
) -> None:
    """Convenience helper to set borders on multiple edges of a cell."""
    sides = {"top": top, "bottom": bottom, "left": left, "right": right}
    for side_name, cfg in sides.items():
        if cfg is not None:
            color = cfg.get("color", "CBD5E1")
            width = cfg.get("width_pt", 1.0)
            style = cfg.get("style", "solid")
            set_cell_border(cell, side_name, color_hex=color, width_pt=width, style=style)


def set_cell_margins(
    cell: _Cell,
    left_inches: float = 0.1,
    right_inches: float = 0.1,
    top_inches: float = 0.08,
    bottom_inches: float = 0.08,
) -> None:
    """Set padding/margins for a table cell in inches.

    1 inch = 914,400 EMUs.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.set("marL", str(int(left_inches * 914400)))
    tcPr.set("marR", str(int(right_inches * 914400)))
    tcPr.set("marT", str(int(top_inches * 914400)))
    tcPr.set("marB", str(int(bottom_inches * 914400)))


def set_cell_shading(cell: _Cell, color_hex: str) -> None:
    """Set background solid fill color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()

    # Remove existing fill elements
    for fill_tag in ["a:solidFill", "a:gradFill", "a:noFill", "a:blipFill"]:
        for old_elem in tcPr.findall(qn(fill_tag)):
            tcPr.remove(old_elem)

    clean_hex = normalize_hex_color(color_hex)
    fill_xml = f'<a:solidFill {nsdecls("a")}><a:srgbClr val="{clean_hex}"/></a:solidFill>'
    fill_element = parse_xml(fill_xml)
    tcPr.append(fill_element)


def set_slide_background(slide: Slide, color_hex: str) -> None:
    """Set solid background color for a slide using OpenXML p:bg.

    Inserts or updates the p:bg element in the slide's p:cSld container.
    """
    clean_hex = normalize_hex_color(color_hex)
    cSld = slide._element.cSld

    # Check for existing background element
    existing_bg = cSld.xpath("p:bg")
    if existing_bg:
        cSld.remove(existing_bg[0])

    bg_xml = (
        f"<p:bg {nsdecls('p', 'a')}>"
        f"  <p:bgPr>"
        f'    <a:solidFill><a:srgbClr val="{clean_hex}"/></a:solidFill>'
        f"    <a:effectLst/>"
        f"  </p:bgPr>"
        f"</p:bg>"
    )
    bg_element = parse_xml(bg_xml)
    # Insert at beginning before p:spTree
    cSld.insert(0, bg_element)

    # Also update python-pptx high-level background fill for internal consistency
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(clean_hex)
    except Exception:
        pass
