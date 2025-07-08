from typing import List, Tuple
from bundle_classes import SKU, Bundle, FILLER_62, FILLER_44
from skyline_packing import fill_remaining_space_with_skyline, find_stackable_skus, _calculate_max_stack_quantity

MAX_LENGTH = 3680

def pack_skus(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """Main entry point for packing SKUs into bundles"""
    # Separate SKUs with bundle override
    override_skus = [sku for sku in skus if sku.data and sku.data.get('Bdl_Override')]
    regular_skus = [sku for sku in skus if sku not in override_skus]

    # Process override bundles first
    override_bundles = _process_override_bundles(override_skus, bundle_width, bundle_height)

    # Group SKUs by color
    color_groups = _group_skus_by_color(regular_skus)
    remaining_color_groups = list(color_groups.items())

    color_bundles = []

    while remaining_color_groups:
        color, color_skus = remaining_color_groups.pop(0)
        color_skus.sort(key=lambda x: x.weight, reverse=True)

        # Pack the initial bundle with this color group
        base_bundles = _pack_skus_with_pattern(color_skus, bundle_width, bundle_height)
        merged_bundles = []

        for base_bundle in base_bundles:
            all_other_groups = remaining_color_groups.copy()
            skus_to_attempt = []
            successful_groups = []

            for other_color, other_skus in all_other_groups:
                test_bundle = Bundle(base_bundle.width, base_bundle.height, base_bundle.max_length)
                for sku in base_bundle.skus:
                    test_bundle.add_sku(sku, sku.x, sku.y, sku.rotated, sku.stacked_quantity)

                test_remaining = fill_remaining_space_with_skyline(test_bundle, other_skus.copy())

                if not test_remaining:
                    base_bundle = test_bundle  # Replace with packed version
                    skus_to_attempt.extend(other_skus)
                    successful_groups.append((other_color, other_skus))

            # Remove successfully packed groups from remaining
            for success_color, _ in successful_groups:
                remaining_color_groups = [grp for grp in remaining_color_groups if grp[0] != success_color]

            base_bundle.resize_to_content()
            _add_filler_material(base_bundle)
            base_bundle.add_packaging()
            merged_bundles.append(base_bundle)

        color_bundles.extend(merged_bundles)

    return override_bundles + color_bundles

def _pack_skus_with_pattern(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """Pack SKUs into bundles using pattern-based algorithm"""
    if not skus:
        return []
    
    skus.sort(key=lambda x: max(x.height, x.width), reverse=True)
    bundles = []
    remaining_skus = skus.copy()

    while remaining_skus:
        # Check if any SKU can fit in a bundle
        if not _can_any_sku_fit(remaining_skus, bundle_width, bundle_height):
            break

        bundle = Bundle(bundle_width, bundle_height, MAX_LENGTH)
        before_count = len(remaining_skus)
        
        remaining_skus = _pack_single_bundle(remaining_skus, bundle)
        
        # Fill remaining space with greedy algorithm
        if remaining_skus and bundle.skus:
            remaining_skus = fill_remaining_space_with_skyline(bundle, remaining_skus)

        # If no progress, skip largest SKU
        if len(remaining_skus) == before_count and remaining_skus:
            largest_sku = max(remaining_skus, key=lambda x: x.width * x.height)
            remaining_skus.remove(largest_sku)

        if bundle.skus:
            bundle.add_packaging()
            bundles.append(bundle)

    return bundles

def _pack_single_bundle(skus: List[SKU], bundle: Bundle) -> List[SKU]:
    """Pack a single bundle using vertical/horizontal pattern"""
    remaining_skus = skus.copy()
    current_y = 0
    is_vertical_row = True
    vertical_row_height = 0
    horizontal_start_y = 0

    # Set bundle max length based on available SKUs
    bundle.max_length = 3680 if max(sku.length for sku in remaining_skus if sku.length) < 3700 else 7340

    # Check for eligible bottom SKUs
    bottom_eligible_skus = [
        sku for sku in remaining_skus
        if sku.can_be_bottom and abs(sku.length - bundle.max_length) <= 100
    ]
    if len(bottom_eligible_skus) > 0:
        # Pack bottom row first if eligible SKUs exist
        bottom_eligible_skus.sort(key=lambda s: max(s.width, s.height), reverse=True)
        if len(bottom_eligible_skus) > 2:
            # If more than 2 eligible SKUs, pack them vertically
            is_vertical_row = True
        else:
            # If 2 or fewer, pack them horizontally
            is_vertical_row = False
        row_height = _place_bottom_row(bundle, bottom_eligible_skus, remaining_skus, is_vertical_row)
        current_y += row_height
        is_vertical_row = not is_vertical_row  # Switch pattern for next row
        # horizontal_start_y = current_y
        # vertical_row_height = row_height

    while remaining_skus and current_y < bundle.height:
        # Pack regular row
        row_height = _pack_row(bundle, remaining_skus, current_y, is_vertical_row, bundle.max_length)

        if row_height == 0:
            break

        current_y += row_height

        # Update pattern for next row
        if is_vertical_row:
            is_vertical_row = False
            vertical_row_height = row_height
            horizontal_start_y = current_y
        else:
            horizontal_section_height = current_y - horizontal_start_y
            if horizontal_section_height >= vertical_row_height:
                is_vertical_row = True
                vertical_row_height = 0
                horizontal_start_y = 0

    return remaining_skus

def _place_bottom_row(bundle: Bundle, bottom_eligible_skus: List[SKU], remaining_skus: List[SKU], is_vertical_row: bool) -> int:
    """Place eligible bottom SKUs horizontally in the first row"""
    bottom_eligible_skus.sort(key=lambda s: max(s.width, s.height), reverse=True)

    current_x = 0
    row_height = 0
    row_skus = []

    for sku in bottom_eligible_skus[:]:
        sku.width, sku.height = _get_sku_dimensions(sku, is_vertical_row)
        if (current_x + sku.width <= bundle.width and
            (row_height == 0 or sku.height <= row_height) and
            _can_fit_in_bundle(sku, current_x, 0, is_vertical_row, bundle)):

            row_skus.append((sku, current_x, 0, is_vertical_row, 1))
            current_x += sku.width
            row_height = max(row_height, sku.height)
            bottom_eligible_skus.remove(sku)
            remaining_skus.remove(sku)

    # Place the row
    for sku, x, y, rotated, qty in row_skus:
        bundle.add_sku(sku, x, y, rotated, qty)

    return row_height

def _pack_row(bundle: Bundle, remaining_skus: List[SKU], current_y: int, is_vertical_row: bool, max_length: int) -> int:
    """Pack a single row of SKUs"""
    row_skus = []
    current_x = 0
    considered_skus = set()

    remaining_skus.sort(key=lambda x: max(x.height, x.width), reverse=True)

    for i, sku in enumerate(remaining_skus):
        # Skip non-eligible SKUs for first row
        # if (current_y == 0 and has_eligible_bottom and 
        #     (not sku.can_be_bottom or abs(sku.length - max_length) > 100)):
        #     continue
            
        if id(sku) in considered_skus:
            continue

        width, height = _get_sku_dimensions(sku, is_vertical_row)

        if (current_x + width <= bundle.width and
            current_y + height <= bundle.height and
            _can_fit_in_bundle(sku, current_x, current_y, is_vertical_row, bundle)):

            # Check support for non-bottom rows
            if current_y != 0 and not _has_sufficient_support(current_x, current_y, width, bundle):
                break

            # Find stackable SKUs
            stackable_skus = find_stackable_skus(sku, remaining_skus, i, max_length)
            stack_quantity = len(stackable_skus) + 1

            # Mark all SKUs in this stack as considered
            considered_skus.add(id(sku))
            for stackable_sku in stackable_skus:
                considered_skus.add(id(stackable_sku))

            row_skus.append((i, sku, current_x, current_y, width, height, stack_quantity, stackable_skus))
            current_x += width

    if not row_skus:
        return 0

    # Place SKUs and remove from remaining list
    skus_to_remove = []
    row_height = 0
    
    for i, sku, x, y, width, height, stack_quantity, stackable_skus in row_skus:
        # Apply rotation if needed
        rotated = _should_rotate_sku(sku, is_vertical_row)
        if rotated:
            sku.width, sku.height = sku.height, sku.width

        bundle.add_sku(sku, x, y, rotated, stack_quantity)
        row_height = max(row_height, height)
        
        skus_to_remove.append(sku)
        skus_to_remove.extend(stackable_skus)

    # Remove placed SKUs
    for sku_to_remove in skus_to_remove:
        if sku_to_remove in remaining_skus:
            remaining_skus.remove(sku_to_remove)

    return row_height

def _add_filler_material(bundle: Bundle) -> None:
    """Add filler material to empty spaces"""
    if not bundle.skus:
        return
    
    fillers = [FILLER_62, FILLER_44]
    
    placed_any = True
    while placed_any:
        placed_any = False
        
        # Generate candidate points
        candidate_points = {(0, 0)}
        for placed_sku in bundle.skus:
            candidate_points.add((placed_sku.x + placed_sku.width, placed_sku.y))
            candidate_points.add((placed_sku.x, placed_sku.y + placed_sku.height))

        # Add grid points for comprehensive coverage
        grid_size = 5
        for x in range(0, int(bundle.width), grid_size):
            for y in range(0, int(bundle.height), grid_size):
                candidate_points.add((x, y))

        # Sort by potential area
        point_priorities = []
        for x, y in candidate_points:
            potential_area = _calculate_potential_area(x, y, bundle)
            point_priorities.append((potential_area, x, y))
        
        point_priorities.sort(key=lambda p: (-p[0], p[2], p[1]))
        
        for _, x, y in point_priorities:
            best_filler, best_config = _find_best_filler(x, y, fillers, bundle)
            
            if best_filler and best_config:
                width, height, rotated = best_config
                stack_quantity = _calculate_max_stack_quantity(best_filler.length, MAX_LENGTH)
                
                filler_copy = SKU(
                    id=best_filler.id,
                    bundleqty=best_filler.bundleqty,
                    width=width,
                    height=height,
                    length=best_filler.length,
                    weight=best_filler.weight,
                    desc=best_filler.desc
                )
                
                bundle.add_sku(filler_copy, x, y, rotated, stack_quantity)
                placed_any = True
                break

# Helper functions
def _group_skus_by_color(skus: List[SKU]) -> dict:
    """Group SKUs by color (text after last period in ID)"""
    color_groups = {}
    for sku in skus:
        color = sku.id.split('.')[-1]
        if color not in color_groups:
            color_groups[color] = []
        color_groups[color].append(sku)
    return color_groups

def _process_override_bundles(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """Process SKUs with bundle override"""
    override_groups = {}
    for sku in skus:
        override_id = sku.data.get('Bdl_Override')
        if override_id not in override_groups:
            override_groups[override_id] = []
        override_groups[override_id].append(sku)

    bundles = []
    for override_skus in override_groups.values():
        bundle = Bundle(bundle_width, bundle_height, MAX_LENGTH)
        for sku in override_skus:
            placed = False
            for x in range(0, bundle_width - int(sku.width), 10):
                for y in range(0, bundle_height - int(sku.height), 10):
                    if _can_fit_in_bundle(sku, x, y, False, bundle):
                        bundle.add_sku(sku, x, y, False)
                        placed = True
                        break
                if placed:
                    break
            if not placed:
                bundles.append(bundle)
                bundle = Bundle(bundle_width, bundle_height, MAX_LENGTH)
                bundle.add_sku(sku, 0, 0, False)
        bundles.append(bundle)
    return bundles

def _can_any_sku_fit(skus: List[SKU], bundle_width: int, bundle_height: int) -> bool:
    """Check if any SKU can fit in a bundle"""
    for sku in skus:
        vert_w, vert_h = _get_sku_dimensions(sku, True)
        horiz_w, horiz_h = _get_sku_dimensions(sku, False)
        if ((vert_w <= bundle_width and vert_h <= bundle_height) or
            (horiz_w <= bundle_width and horiz_h <= bundle_height)):
            return True
    return False

def _can_place_sku_at_position(sku: SKU, x: int, y: int, width: int, height: int, bundle: Bundle) -> bool:
    """Check if SKU can be placed at specific position with given dimensions"""
    if x + width > bundle.width or y + height > bundle.height:
        return False
    
    for placed_sku in bundle.skus:
        if (x < placed_sku.x + placed_sku.width and
            x + width > placed_sku.x and
            y < placed_sku.y + placed_sku.height and
            y + height > placed_sku.y):
            return False
    return True


def _has_sufficient_support(x: int, y: int, width: int, bundle: Bundle, threshold: float = 0.7) -> bool:
    """Check if position has sufficient support from SKUs below"""
    support_segments = []
    for sku in bundle.skus:
        if sku.y + sku.height == y:
            overlap_start = max(x, sku.x)
            overlap_end = min(x + width, sku.x + sku.width)
            if overlap_end > overlap_start:
                support_segments.append((overlap_start, overlap_end))

    total_supported_width = sum(end - start for start, end in support_segments)
    return (total_supported_width / width) >= threshold

def _get_sku_dimensions(sku: SKU, vertical: bool) -> Tuple[int, int]:
    """Get SKU dimensions based on orientation"""
    if vertical:
        return min(sku.width, sku.height), max(sku.width, sku.height)
    else:
        return max(sku.width, sku.height), min(sku.width, sku.height)

def _can_fit_in_bundle(sku: SKU, x: int, y: int, vertical: bool, bundle: Bundle) -> bool:
    """Check if SKU can fit at position with given orientation"""
    width, height = _get_sku_dimensions(sku, vertical)
    return _can_place_sku_at_position(sku, x, y, width, height, bundle)

def _should_rotate_sku(sku: SKU, is_vertical_row: bool) -> bool:
    """Determine if SKU should be rotated based on row orientation"""
    if is_vertical_row:
        return sku.width > sku.height
    else:
        return sku.width < sku.height

def _calculate_potential_area(x: int, y: int, bundle: Bundle) -> int:
    """Calculate potential area available at a position"""
    max_width = bundle.width - x
    max_height = bundle.height - y
    
    for placed_sku in bundle.skus:
        if placed_sku.x >= x and placed_sku.y >= y:
            if placed_sku.x < x + max_width:
                max_width = min(max_width, placed_sku.x - x)
            if placed_sku.y < y + max_height:
                max_height = min(max_height, placed_sku.y - y)
    
    return max_width * max_height

def _find_best_filler(x: int, y: int, fillers: List[SKU], bundle: Bundle) -> Tuple[SKU, Tuple[int, int, bool]]:
    """Find best filler for a position"""
    best_filler = None
    best_config = None
    best_area = 0
    
    for filler in fillers:
        orientations = [
            (filler.width, filler.height, False),
            (filler.height, filler.width, True)
        ]
        
        for width, height, rotated in orientations:
            if _can_place_sku_at_position(filler, x, y, width, height, bundle):
                if y != 0 and not _has_sufficient_support(x, y, width, bundle):
                    continue
                
                area = width * height
                if area > best_area:
                    best_area = area
                    best_filler = filler
                    best_config = (width, height, rotated)
    
    return best_filler, best_config

