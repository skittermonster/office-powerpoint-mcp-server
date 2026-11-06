"""Geometry, layout, coordinate, and typography sizing helpers."""

from typing import Any


def calculate_font_size(
    text: str,
    container_width_inches: float,
    container_height_inches: float,
    role: str = "body",
    min_size: float = 8.0,
    max_size: float = 40.0,
) -> float:
    """Calculate an optimal font size (in pt) based on text length and box dimensions.

    Prevents text overflow by scaling font down as character count increases.
    """
    char_count = max(1, len(text.strip()))

    if role == "title":
        target_max = min(max_size, 40.0)
        target_min = max(min_size, 18.0)
        # Approximate: title line count
        if char_count < 30:
            size = target_max
        elif char_count < 60:
            size = target_max * 0.85
        elif char_count < 100:
            size = target_max * 0.70
        else:
            size = target_min
    elif role == "subtitle":
        target_max = min(max_size, 22.0)
        target_min = max(min_size, 12.0)
        if char_count < 50:
            size = target_max
        elif char_count < 120:
            size = target_max * 0.85
        else:
            size = target_min
    elif role == "kpi":
        target_max = min(max_size, 44.0)
        target_min = max(min_size, 22.0)
        if char_count < 8:
            size = target_max
        elif char_count < 15:
            size = target_max * 0.8
        else:
            size = target_min
    elif role == "header":
        target_max = min(max_size, 13.0)
        target_min = max(min_size, 9.0)
        if char_count < 15:
            size = target_max
        elif char_count < 35:
            size = target_max * 0.9
        else:
            size = target_min
    else:  # body / cell
        target_max = min(max_size, 14.0)
        target_min = max(min_size, 8.0)
        if char_count < 80:
            size = target_max
        elif char_count < 200:
            size = 11.0
        elif char_count < 400:
            size = 9.5
        else:
            size = target_min

    # Also constrain by container area if dimensions are small
    area = container_width_inches * container_height_inches
    if area < 1.0:
        size = min(size, 10.0)
    elif area < 3.0:
        size = min(size, 14.0)

    return max(min_size, min(max_size, round(size, 1)))


def fit_image_in_bounding_box(
    img_width: int | float,
    img_height: int | float,
    left: float,
    top: float,
    max_width: float,
    max_height: float,
    preserve_aspect_ratio: bool = True,
    center: bool = True,
) -> tuple[float, float, float, float]:
    """Calculate placed (left, top, width, height) in inches fitting within a bounding box.

    Args:
        img_width: Raw image pixel or point width
        img_height: Raw image pixel or point height
        left: Target bounding box left in inches
        top: Target bounding box top in inches
        max_width: Target bounding box max width in inches
        max_height: Target bounding box max height in inches
        preserve_aspect_ratio: Whether to maintain aspect ratio
        center: Whether to center image in bounding box

    Returns:
        (placed_left, placed_top, placed_width, placed_height) in inches
    """
    if not preserve_aspect_ratio or img_width <= 0 or img_height <= 0:
        return left, top, max_width, max_height

    aspect = img_width / img_height
    box_aspect = max_width / max_height

    if aspect > box_aspect:
        # Constrained by width
        placed_w = max_width
        placed_h = max_width / aspect
    else:
        # Constrained by height
        placed_h = max_height
        placed_w = max_height * aspect

    placed_l = left + ((max_width - placed_w) / 2.0 if center else 0.0)
    placed_t = top + ((max_height - placed_h) / 2.0 if center else 0.0)

    return round(placed_l, 4), round(placed_t, 4), round(placed_w, 4), round(placed_h, 4)


def calculate_column_widths(
    total_width: float,
    headers: list[str],
    sample_rows: list[list[Any]] | None = None,
    min_col_width: float = 0.8,
) -> list[float]:
    """Calculate professional column widths for financial/data tables.

    Column 0 (labels/metrics) is typically wider than numeric value columns.
    """
    col_count = len(headers)
    if col_count == 0:
        return []
    if col_count == 1:
        return [total_width]

    # Calculate rough text lengths for each column
    max_lens = [len(str(h)) for h in headers]
    if sample_rows:
        for row in sample_rows[:10]:
            for i, val in enumerate(row):
                if i < col_count:
                    max_lens[i] = max(max_lens[i], len(str(val)))

    # First column typically holds longer category/metric labels
    # Weights based on character lengths with floor
    weights = [max(12, max_lens[0] * 1.2)]
    for length_val in max_lens[1:]:
        weights.append(max(8, float(length_val)))

    total_weight = sum(weights)
    widths = [(w / total_weight) * total_width for w in weights]

    # Enforce minimum column widths
    for i in range(len(widths)):
        if widths[i] < min_col_width:
            widths[i] = min_col_width

    # Normalize to exactly match total_width
    current_sum = sum(widths)
    scale = total_width / current_sum
    normalized_widths = [round(w * scale, 3) for w in widths]

    # Adjust rounding discrepancy on last column
    diff = total_width - sum(normalized_widths)
    normalized_widths[-1] = round(normalized_widths[-1] + diff, 3)

    return normalized_widths
