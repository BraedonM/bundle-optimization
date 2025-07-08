from typing import List, Dict, Tuple, Optional
from collections import defaultdict
from math import ceil
from bundle_classes import SKU, Bundle, FILLER_62, FILLER_44

class PackingError(Exception):
    pass

def pack_skus(skus: List[SKU], maxWidth: float, maxHeight: float) -> List[Bundle]:
    # Group SKUs by override and color
    override_groups = defaultdict(list)
    non_override_skus = []
    
    for sku in skus:
        if sku.data and sku.data.get('Bdl_Override') not in [None, '']:
            override_groups[sku.data['Bdl_Override']].append(sku)
        else:
            non_override_skus.append(sku)
    
    # Group non-override SKUs by color
    color_groups = defaultdict(list)
    for sku in non_override_skus:
        color = sku.id.split('.')[-1]
        color_groups[color].append(sku)
    
    # Combine all groups
    groups = list(override_groups.values()) + list(color_groups.values())
    
    bundles = []
    for group in groups:
        while group:
            # remove SKUs that are already in bundles
            group_bundles, group = pack_group(group, maxWidth, maxHeight)
            bundles.extend(group_bundles)
    
    return bundles

def pack_group(group: List[SKU], maxWidth: float, maxHeight: float) -> List[Bundle]:
    if not group:
        return [], []
    
    # Determine bundle length
    max_sku_length = max(sku.length for sku in group)
    if max_sku_length <= 3700:
        bundle_length = 3680
        bottom_range = (3600, 3700)
    else:
        bundle_length = 7340
        bottom_range = (7300, 7400)
    
    # Separate bottom-eligible SKUs
    bottom_eligible = [
        sku for sku in group 
        if sku.can_be_bottom and bottom_range[0] <= sku.length <= bottom_range[1]
    ]
    
    if not bottom_eligible:
        bottom_eligible = [
            sku for sku in group
        ]
    
    # Create bundle and pack bottom row
    bundle = Bundle(0, 0, bundle_length)
    current_x = 0
    bottom_row_height = 0
    bottom_row_items = []
    
    bottom_couldnt_fit = []
    
    # Pack bottom row first
    for sku in sorted(bottom_eligible, key=lambda s: s.width, reverse=True):
        placed = False
        if len(bottom_row_items) > 2:
            # Try rotated orientation
            if current_x + sku.height <= maxWidth:
                placed_sku = bundle.add_sku(sku, current_x, 0, True)
                bottom_row_items.append(placed_sku)
                current_x += sku.height
                bottom_row_height = max(bottom_row_height, sku.width)
                placed = True
            # Try natural orientation
            elif current_x + sku.width <= maxWidth:
                placed_sku = bundle.add_sku(sku, current_x, 0, False)
                bottom_row_items.append(placed_sku)
                current_x += sku.width
                bottom_row_height = max(bottom_row_height, sku.height)
                placed = True
        else:
            # Try natural orientation
            if current_x + sku.width <= maxWidth:
                placed_sku = bundle.add_sku(sku, current_x, 0, False)
                bottom_row_items.append(placed_sku)
                current_x += sku.width
                bottom_row_height = max(bottom_row_height, sku.height)
                placed = True
            # Try rotated orientation
            elif current_x + sku.height <= maxWidth:
                placed_sku = bundle.add_sku(sku, current_x, 0, True)
                bottom_row_items.append(placed_sku)
                current_x += sku.height
                bottom_row_height = max(bottom_row_height, sku.width)
                placed = True

        if not placed:
            bottom_couldnt_fit.append(sku)

    remaining = [sku for sku in group if sku not in bottom_eligible] + bottom_couldnt_fit
    remaining.sort(key=lambda s: max(s.height, s.width), reverse=True)

    skus_couldnt_fit = []

    # Pack rest of SKUs using skyline algorithm
    count_vertical = 0
    count_horizontal = 0
    while remaining:
        profile = get_top_surface_profile(bundle, maxWidth)
        lowest_points = find_lowest_points(profile)

        placed_any = False

        for x, y in lowest_points:
            for i, sku in enumerate(remaining):
                if count_vertical < count_horizontal:
                    orientations = [ # Try to place vertically first
                        (sku.height, sku.width, True),
                        (sku.width, sku.height, False)
                    ]
                else:
                    orientations = [ # Try to place horizontally first
                        (sku.width, sku.height, False),
                        (sku.height, sku.width, True)
                    ]
                for w, h, rotated in orientations:
                    if x + w > maxWidth or y + h > maxHeight:
                        continue
                    if has_sufficient_support(x, y, w, bundle) and check_overlap(x, y, w, h, bundle):
                        # snug_x = find_snug_x_left(x, y, w, h, bundle, rotated)
                        bundle.add_sku(sku, x, y, rotated)
                        del remaining[i]
                        placed_any = True
                        if rotated:
                            count_vertical += 1
                        else:
                            count_horizontal += 1
                        break
                if placed_any:
                    break
            if placed_any:
                break

        if not placed_any:
            # No more placements possible, stop the loop
            skus_couldnt_fit.extend(remaining)
            break
    
    # Resize bundle to actual content
    bundle.resize_to_content()
    
    # Add packaging materials
    bundle.add_packaging()
    
    return [bundle], skus_couldnt_fit

def has_sufficient_support(x, y, width, bundle: Bundle, threshold=0.7, max_gap=10, ret_real_val=False) -> bool:
    """
    Check if at least `threshold` fraction of the bottom width is supported by SKUs up to `max_gap` mm below.
    """
    support_segments = []
    support_range_start = y - max_gap
    support_range_end = y

    for sku in bundle.skus:
        sku_top = sku.y + sku.height
        # If the top of the supporting SKU is within the acceptable vertical support range
        if support_range_start <= sku_top <= support_range_end:
            # Compute horizontal overlap
            overlap_x_start = max(x, sku.x)
            overlap_x_end = min(x + width, sku.x + sku.width)
            if overlap_x_end > overlap_x_start:
                support_segments.append((overlap_x_start, overlap_x_end))

    # Merge overlapping support segments to calculate total unique support length
    support_segments.sort()
    merged_segments = []
    for start, end in support_segments:
        if not merged_segments or start > merged_segments[-1][1]:
            merged_segments.append((start, end))
        else:
            merged_segments[-1] = (merged_segments[-1][0], max(merged_segments[-1][1], end))

    total_supported_width = sum(end - start for start, end in merged_segments)
    if ret_real_val:
        return (total_supported_width / width)
    return (total_supported_width / width) >= threshold

def get_top_surface_profile(bundle: Bundle, max_width: float) -> List[Tuple[float, float]]:
    """
    Returns a list of tuples (x, y) representing the top surface profile at intervals along x.
    """
    resolution = 1
    surface = []

    for x in range(0, int(max_width), resolution):
        max_y = 0
        for sku in bundle.skus:
            if sku.x <= x < sku.x + sku.width:
                max_y = max(max_y, sku.y + sku.height)
        surface.append((x, max_y))

    surface = [(x, y) for x, y in surface if y > 0]
    return surface

def find_lowest_points(profile: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    """
    Sorts points by lowest y value (height) ascending.
    """
    return sorted(profile, key=lambda p: p[1])

def find_snug_x_left(x: float, y: float, w: float, h: float, bundle: Bundle, rotated: bool) -> float:
    """
    Slide the SKU left until it either hits the left wall or another SKU, or loses sufficient support.
    """
    while x > 0:
        next_x = x - 1
        if not has_sufficient_support(next_x, y, w, bundle) or (
            has_sufficient_support(next_x, y, w, bundle, ret_real_val=True) < has_sufficient_support(x, y, w, bundle, ret_real_val=True)):
            break
        # Check for overlap with existing SKUs
        if not check_overlap(next_x, y, w, h, bundle):
            break
        x = next_x
    return x

def check_overlap(x: float, y: float, w: float, h: float, bundle: Bundle) -> bool:
    """
    Check if the SKU at (x, y) with dimensions (w, h) overlaps with any existing SKUs in the bundle.
    Returns True if NO overlap (safe to place), False if overlap detected.
    """
    for sku in bundle.skus:
        # Get the actual dimensions of the existing SKU based on its rotation
        if sku.rotated:
            sku_w, sku_h = sku.height, sku.width
        else:
            sku_w, sku_h = sku.width, sku.height
        
        # Define boundaries of existing SKU
        sku_left = sku.x
        sku_right = sku.x + sku_w
        sku_bottom = sku.y
        sku_top = sku.y + sku_h
        
        # Define boundaries of new SKU
        new_left = x
        new_right = x + w
        new_bottom = y
        new_top = y + h
        
        # Check for overlap using standard rectangle overlap detection
        # Two rectangles overlap if they overlap in both x and y dimensions
        x_overlap = not (new_right <= sku_left or new_left >= sku_right)
        y_overlap = not (new_top <= sku_bottom or new_bottom >= sku_top)
        
        if x_overlap and y_overlap:
            return False  # Overlap detected, not safe to place
    
    return True  # No overlap found, safe to place
