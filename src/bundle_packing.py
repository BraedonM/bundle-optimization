from typing import List, Tuple
from bundle_classes import SKU, Bundle, FILLER_62, FILLER_44

MAX_LENGTH = 3680
BOTTOM_ROW_LENGTH = 0
REMOVED_SKUS = []

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

    can_try_merge_bundles = []
    final_bundles = []

    while remaining_color_groups:
        color, color_skus = remaining_color_groups.pop(0)
        color_skus.sort(key=lambda x: max(x.width, x.height), reverse=True)

        base_bundles = _pack_skus_with_pattern(color_skus, bundle_width, bundle_height)
        can_try_merge_bundles.extend(base_bundles)

    # try to merge bundles if they can all fit in one
    merged_bundles = _try_merge_bundles(can_try_merge_bundles, bundle_width, bundle_height)
    final_bundles.extend(merged_bundles)

    # remove any empty bundles
    final_bundles = [bundle for bundle in final_bundles if (bundle.width > 0 and bundle.height > 0)]
    for bundle in final_bundles:
        bundle.add_packaging()  # Add packaging to each bundle

    return override_bundles + final_bundles, REMOVED_SKUS

def _try_merge_bundles(bundles: List[Bundle], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """Attempt to merge bundles if they can all fit in one bundle"""
    best_bundles = []
    mid_bundles = []
    bad_bundles = []
    # goal is to merge bundles with bottom-eligible SKUs first, prioritizing full-length
    for bundle in bundles:
        bottom_able_skus = [sku for sku in bundle.skus if sku.can_be_bottom]
        if bottom_able_skus:
            if any([sku.length > 7000 for sku in bottom_able_skus]):
                best_bundles.append(bundle)
            else:
                mid_bundles.append(bundle)
        else:
            bad_bundles.append(bundle)
    # sort bundles by size
    bundles = best_bundles + mid_bundles + bad_bundles

    merging_able = True
    while merging_able:
         # pick two bundles and try to merge them
        merging_able = False
        for i in range(len(bundles)):
            for j in range(i + 1, len(bundles)):
                bundle1 = bundles[i]
                bundle2 = bundles[j]

                # get all SKUs from both bundles
                all_skus = bundle1.skus + bundle2.skus
                # deconstruct PlacedSKU objects into just SKU objects for consistency
                regular_skus = []
                for sku in all_skus:
                    regular_skus.append(SKU(id=sku.id, bundleqty=sku.bundleqty, width=sku.width,
                        height=sku.height, length=sku.length, weight=sku.weight, desc=sku.desc,
                        can_be_bottom=sku.can_be_bottom, data=sku.data))
                all_skus = regular_skus
                # try to pack them into a new bundle
                merged_bundles = _pack_skus_with_pattern(all_skus, bundle_width, bundle_height)
                if len(merged_bundles) == 1:
                    # if they fit into one bundle, remove the original bundles
                    bundles.pop(j)
                    bundles.pop(i)
                    bundles.append(merged_bundles[0])
                    merging_able = True
                    break
            if merging_able:
                break
    # After merging, add filler material to each bundle
    for bundle in bundles:
        bundle.resize_to_content()
        _add_filler_material(bundle)

    return bundles

def _pack_skus_with_pattern(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """Pack SKUs into bundles using pattern-based algorithm"""
    global REMOVED_SKUS
    if not skus:
        return []
    skus.sort(key=lambda x: max(x.height, x.width), reverse=True)
    
    bundles = []
    remaining_skus = skus.copy()

    while remaining_skus:
        # Check if any SKU can fit in a bundle
        if not any(sku.can_be_bottom for sku in remaining_skus):
            # If no SKU can be bottom, set all to True
            for sku in remaining_skus:
                sku.can_be_bottom = True
        if not _can_any_sku_fit(remaining_skus, bundle_width, bundle_height):
            break

        bundle = Bundle(bundle_width, bundle_height, MAX_LENGTH)
        before_count = len(remaining_skus)
        
        remaining_skus = _pack_single_bundle(remaining_skus, bundle)
        
        # If no progress, skip largest SKU
        if len(remaining_skus) == before_count and remaining_skus:
            largest_sku = max(remaining_skus, key=lambda x: x.width * x.height)
            if largest_sku not in [sku.id for sku in REMOVED_SKUS]:
                REMOVED_SKUS.append(largest_sku)
            remaining_skus.remove(largest_sku)

        if bundle.skus:
            bundles.append(bundle)

    return bundles

def _pack_single_bundle(skus: List[SKU], bundle: Bundle) -> List[SKU]:
    """Pack a single bundle using vertical/horizontal pattern"""
    remaining_skus = skus.copy()
    current_y = 0
    is_vertical_row = True

    # Set bundle max length based on available SKUs
    bundle.max_length = 3680 if max(sku.length for sku in remaining_skus if sku.length) < 3700 else 7340

    # Check for eligible bottom SKUs
    bottom_eligible_skus = [
        sku for sku in remaining_skus
        if sku.can_be_bottom and abs(sku.length - bundle.max_length) <= 100
    ]
    if len(bottom_eligible_skus) > 0:
        # Pack bottom row first if eligible SKUs exist
        row_height = _place_bottom_row(bundle, bottom_eligible_skus, remaining_skus, is_vertical_row)
        current_y += row_height
        is_vertical_row = False

    short_skus = [sku for sku in remaining_skus if sku.length <= 609]
    remaining_skus = [sku for sku in remaining_skus if sku.length > 609]

    while remaining_skus and current_y < bundle.height:
        # Pack regular row
        row_height = _pack_row(bundle, remaining_skus, current_y, is_vertical_row, bundle.max_length)

        if row_height == 0:
            break

        current_y += row_height
        remaining_skus = fill_row_greedy(bundle, remaining_skus, current_y)
        if is_vertical_row:
            is_vertical_row = False

    remaining_skus = fill_remaining_greedy(bundle, remaining_skus)

    if current_y == 0:
        is_vertical_row = True
    if short_skus and current_y < bundle.height:
        while short_skus and current_y < bundle.height:
            # Pack short SKUs in a greedy manner
            row_height = _pack_row(bundle, short_skus, current_y, is_vertical_row, bundle.max_length)
            if row_height == 0:
                break

            current_y += row_height
            short_skus = fill_row_greedy(bundle, short_skus, current_y)

            # Update pattern for next row
            if is_vertical_row:
                is_vertical_row = False

        # Fill remaining gaps after initial packing
        short_remaining_skus = fill_remaining_greedy(bundle, short_skus)
        remaining_skus += short_remaining_skus

    return remaining_skus

def _place_bottom_row(bundle: Bundle, bottom_eligible_skus: List[SKU], remaining_skus: List[SKU], is_vertical_row: bool) -> int:
    """Place eligible bottom SKUs horizontally in the first row"""
    global BOTTOM_ROW_LENGTH
    for sku in bottom_eligible_skus:
        sku.width, sku.height = _get_sku_dimensions(sku, is_vertical_row)
    bottom_eligible_skus.sort(key=lambda s: s.height, reverse=True)

    current_x = 0
    row_height = 0
    row_skus = []

    for i, sku in enumerate(bottom_eligible_skus[:]):
        if (current_x + sku.width <= bundle.width and
            (row_height == 0 or sku.height <= row_height) and
            _can_fit_in_bundle(sku, current_x, 0, is_vertical_row, bundle)
            and _sku_within_height_range(sku, row_skus)):

            row_skus.append((i, sku, current_x, 0, is_vertical_row))
            current_x += sku.width
            row_height = max(row_height, sku.height)
            bottom_eligible_skus.remove(sku)
            remaining_skus.remove(sku)
            BOTTOM_ROW_LENGTH = current_x

    # Place the row
    for i, sku, x, y, rotated in row_skus:
        bundle.add_sku(sku, x, y, rotated)

    return row_height

def _pack_row(bundle: Bundle, remaining_skus: List[SKU], current_y: int, is_vertical_row: bool, max_length: int) -> int:
    """Pack a single row of SKUs"""
    row_skus = []
    current_x = 0
    considered_skus = set()
    
    unique_skus = []
    for sku in remaining_skus:
        if sku.id not in [s.id for s in unique_skus]:
            unique_skus.append(sku)
    
    # if len(unique_skus) <= 2 and is_vertical_row and all(max(sku.width, sku.height) > 2 * min(sku.width, sku.height) for sku in unique_skus):
    #     is_vertical_row = False  # Force horizontal if only a couple SKUs that are very tall
    for sku in remaining_skus:
        sku.width, sku.height = _get_sku_dimensions(sku, is_vertical_row)
    remaining_skus.sort(key=lambda s: s.height, reverse=True)

    for i, sku in enumerate(remaining_skus):
        width, height = _get_sku_dimensions(sku, is_vertical_row)
        if (
            # sku has been placed already
            id(sku) in considered_skus or
            # sku doesn't have valid dimensions
            width <= 0 or height <= 0 or
            # sku is outside +-25mm height of other SKUs in the row
            not _sku_within_height_range(sku, row_skus)
        ):
            continue

        if (current_x + width <= bundle.width and
            current_y + height <= bundle.height and (current_y == 0 or current_y + height <= BOTTOM_ROW_LENGTH) and
            _can_fit_in_bundle(sku, current_x, current_y, is_vertical_row, bundle)):

            # Check support for non-bottom rows
            if current_y != 0 and not _has_sufficient_support(current_x, current_y, width, bundle):
                continue

            # Find stackable SKUs
            stackable_skus = find_stackable_skus(sku, remaining_skus, considered_skus, i, max_length)
            
            # Mark all SKUs in this stack as considered
            considered_skus.add(id(sku))
            for stackable_sku in stackable_skus:
                considered_skus.add(id(stackable_sku))

            row_skus.append((i, sku, current_x, current_y, width, height, stackable_skus))
            current_x += width
    if not row_skus:
        return 0

    # Place SKUs and remove from remaining list
    skus_to_remove = []
    row_height = 0
    
    for i, sku, x, y, width, height, stackable_skus in row_skus:
        # Apply rotation if needed for target SKU
        rotated = _should_rotate_sku(sku, is_vertical_row)
        if rotated:
            sku.width, sku.height = sku.height, sku.width
        current_max_height = sku.height

        # Place target SKU
        bundle.add_sku(sku, x, y, rotated)
        
        # Place stackable SKUs sorted by length (largest first)
        stackable_skus_sorted = sorted(stackable_skus, key=lambda s: s.length, reverse=True)
        for stack_sku in stackable_skus_sorted:
            stack_rotated = _should_rotate_sku(stack_sku, is_vertical_row)
            if stack_rotated:
                stack_sku.width, stack_sku.height = stack_sku.height, stack_sku.width
                
            # Update max height for this stack
            if stack_sku.height > current_max_height:
                current_max_height = stack_sku.height
                
            # Place stack SKU
            bundle.add_sku(stack_sku, x, y, stack_rotated)
            
        row_height = max(row_height, current_max_height)
        
        skus_to_remove.append(sku)
        skus_to_remove.extend(stackable_skus)

    # Remove placed SKUs
    for sku_to_remove in skus_to_remove:
        if sku_to_remove in remaining_skus:
            remaining_skus.remove(sku_to_remove)

    return row_height

def _add_filler_material(bundle: Bundle) -> None:
    """Add filler material to empty spaces, avoiding edges when possible"""
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

        # Sort by potential area and interior priority
        point_priorities = []
        for x, y in candidate_points:
            potential_area = _calculate_potential_area(x, y, bundle)
            
            # Calculate distance to nearest edge
            dist_left = x
            dist_right = bundle.width - x
            dist_top = bundle.height - y
            dist_bottom = y
            min_dist = min(dist_left, dist_right, dist_top, dist_bottom)
            
            # Prioritize interior points (min_dist > 50mm gets bonus)
            interior_bonus = 1.0
            if min_dist > 50:
                interior_bonus = 2.0  # Double priority for interior points
                
            point_priorities.append((potential_area * interior_bonus, min_dist, x, y))
        
        # Sort by potential area (with bonus) then by distance to edge
        point_priorities.sort(key=lambda p: (-p[0], -p[1]))
        
        for _, min_dist, x, y in point_priorities:
            # Skip bottom row for filler
            if y == 0:
                continue
                
            best_filler, best_config = _find_best_filler(x, y, fillers, bundle)
            
            if best_filler and best_config:
                width, height, rotated = best_config
                
                filler_copy = SKU(
                    id=best_filler.id,
                    bundleqty=best_filler.bundleqty,
                    width=width,
                    height=height,
                    length=best_filler.length,
                    weight=best_filler.weight,
                    desc=best_filler.desc
                )
                
                bundle.add_sku(filler_copy, x, y, rotated)
                placed_any = True
                break

# Helper functions
def _group_skus_by_color(skus: List[SKU]) -> dict:
    """Group SKUs by color (text after last period in ID)"""
    color_groups = {}
    for sku in skus:
        color = sku.id.split('.')[-1]
        if "Partial" in color:
            color = color.replace("_Partial", "")
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
    if y == 0 and (not sku.can_be_bottom or (bundle.max_length == 7340 and sku.length < 3700)):
        return False

    for placed_sku in bundle.skus:
        if (x < placed_sku.x + placed_sku.width and
            x + width > placed_sku.x and
            y < placed_sku.y + placed_sku.height and
            y + height > placed_sku.y):
            return False
    return True

def _has_sufficient_support(x: int, y: int, width: int, bundle: Bundle, threshold: float = 0.85, get_value: bool = False) -> bool:
    """Check if position has sufficient support from SKUs below"""
    buffer = 5
    support_segments = []
    # Loop through SKUs to find overlaps
    for sku in bundle.skus:
        # If SKU is too short, skip it
        if sku.length <= 609:
            continue
        # vertical tolerance of +-5mm for support
        if y - buffer <= sku.y + sku.height <= y + buffer:
            overlap_start = max(x, sku.x)
            overlap_end = min(x + width, sku.x + sku.width)
            if overlap_end > overlap_start:
                support_segments = list(set(support_segments).union(list(range(round(overlap_start), round(overlap_end)))))

    total_supported_width = len(support_segments)
    if get_value:
        return total_supported_width / width
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
    """Find best filler for a position, avoiding edges when possible"""
    best_filler = None
    best_config = None
    best_area = 0
    min_edge_distance = 0
    
    for filler in fillers:
        orientations = [
            (filler.width, filler.height, False),
            (filler.height, filler.width, True)
        ]
        
        for width, height, rotated in orientations:
            if not _can_place_sku_at_position(filler, x, y, width, height, bundle):
                continue
                
            if y != 0 and not _has_sufficient_support(x, y, width, bundle):
                continue
            
            # Calculate distance to nearest edge
            dist_left = x
            dist_right = bundle.width - (x + width)
            dist_top = bundle.height - (y + height)
            dist_bottom = y
            current_min_dist = min(dist_left, dist_right, dist_top, dist_bottom)
            
            area = width * height
            # Prefer fillers that are farther from edges
            if (current_min_dist > min_edge_distance or 
                (current_min_dist == min_edge_distance and area > best_area)):
                best_area = area
                best_filler = filler
                best_config = (width, height, rotated)
                min_edge_distance = current_min_dist
    
    return best_filler, best_config

def fill_remaining_greedy(bundle: Bundle, remaining_skus: List[SKU]) -> List[SKU]:
    """Fill remaining gaps in bundle with greedy placement approach, avoiding filler on edges"""
    if not remaining_skus:
        return remaining_skus

    placed_any = True
    while placed_any and remaining_skus:
        placed_any = False
        # Generate candidate points (existing corners + grid points)
        candidate_points = set()
        # Add corners of existing SKUs
        for sku in bundle.skus:
            candidate_points.add((sku.x + sku.width, sku.y))  # Right of existing SKU
            candidate_points.add((sku.x, sku.y + sku.height))  # Below existing SKU
        # Add grid points for dense coverage
        grid_size = 25  # 50mm grid for performance
        for x in range(0, bundle.width, grid_size):
            for y in range(0, bundle.height, grid_size):
                candidate_points.add((x, y))
        candidate_points = sorted(candidate_points, key=lambda p: (p[1], p[0]))  # Sort by y then x

        # Try largest SKUs first
        remaining_skus.sort(key=lambda s: s.width * s.height, reverse=True)

        # Track which SKUs we've already considered for stacking
        considered_skus = set()

        for i, sku in enumerate(remaining_skus[:]):  # Iterate over copy
            if id(sku) in considered_skus:
                continue
                
            for rotated in [False, True]:
                w, h = _get_sku_dimensions(sku, rotated)

                # Skip if too big for bundle
                if w > bundle.width or h > bundle.height:
                    continue

                for (x, y) in candidate_points:
                    if x + w > bundle.width or y + h > bundle.height or y + h > BOTTOM_ROW_LENGTH:
                        continue

                    if _can_place_sku_at_position(sku, x, y, w, h, bundle):
                        # Check support if not on bottom
                        if y > 0 and not _has_sufficient_support(x, y, w, bundle) or (
                            y == 0 and (abs(sku.length - bundle.max_length) > 100 or not sku.can_be_bottom)):
                            continue

                        # Find stackable SKUs
                        stackable_skus = find_stackable_skus(sku, remaining_skus, considered_skus, i, bundle.max_length)
                        
                        # Mark all SKUs in this stack as considered
                        considered_skus.add(id(sku))
                        for stackable_sku in stackable_skus:
                            considered_skus.add(id(stackable_sku))

                        # Place the main SKU
                        if rotated:
                            sku.width, sku.height = sku.height, sku.width
                        bundle.add_sku(sku, x, y, rotated)
                        
                        # Place stackable SKUs at the same position
                        for stack_sku in stackable_skus:
                            stack_rotated = rotated  # Use same rotation as main SKU
                            if stack_rotated:
                                stack_sku.width, stack_sku.height = stack_sku.height, stack_sku.width
                            bundle.add_sku(stack_sku, x, y, stack_rotated)
                        
                        # Remove all placed SKUs from remaining list
                        remaining_skus.remove(sku)
                        for stack_sku in stackable_skus:
                            if stack_sku in remaining_skus:
                                remaining_skus.remove(stack_sku)
                        
                        placed_any = True
                        break  # Break candidate points loop
                if placed_any:
                    break  # Break rotation loop
            if placed_any:
                break  # Break SKU loop to restart with new candidate points
                
    return remaining_skus

def fill_row_greedy(bundle: Bundle,
                    remaining_skus: List[SKU],
                    y_limit: int) -> List[SKU]:
    """
    Try to place any remaining SKU into the area y ∈ [0, y_limit),
    using the same greedy corner/stack logic as fill_remaining_greedy,
    but *never* placing anything with y >= y_limit.
    """
    placed = True
    while placed and remaining_skus:
        placed = False
        # same candidate‐point generation...
        candidate_points = set()
        for sku in bundle.skus:
            candidate_points.add((sku.x + sku.width, sku.y))
            candidate_points.add((sku.x, sku.y + sku.height))
        # grid
        for gx in range(0, bundle.width, 50):
            for gy in range(0, round(y_limit), 50):
                candidate_points.add((gx, gy))
        
        # Track which SKUs we've already considered for stacking
        considered_skus = set()
        
        for i, sku in enumerate(sorted(remaining_skus, key=lambda s: s.width*s.height, reverse=True)):
            if id(sku) in considered_skus:
                continue
                
            for rot in (False, True):
                w, h = _get_sku_dimensions(sku, rot)
                if w > bundle.width or h > bundle.height or h > y_limit:
                    continue
                for x,y in sorted(candidate_points, key=lambda p: (p[1], p[0])):
                    if x+w > bundle.width or y+h > y_limit or (y == 0 and abs(sku.length - bundle.max_length) > 100):
                        continue
                    if _can_place_sku_at_position(sku, x, y, w, h, bundle) and \
                        ((y == 0 and sku.can_be_bottom) or _has_sufficient_support(x, y, w, bundle)):
                        
                        # Find stackable SKUs
                        original_index = remaining_skus.index(sku)
                        stackable_skus = find_stackable_skus(sku, remaining_skus, considered_skus, original_index, bundle.max_length)
                        
                        # Mark all SKUs in this stack as considered
                        considered_skus.add(id(sku))
                        for stackable_sku in stackable_skus:
                            considered_skus.add(id(stackable_sku))
                        
                        # Place the main SKU
                        if _should_rotate_sku(sku, rot):
                            sku.width, sku.height = sku.height, sku.width
                        bundle.add_sku(sku, x, y, rot)
                        
                        # Place stackable SKUs at the same position
                        for stack_sku in stackable_skus:
                            stack_rotated = rot  # Use same rotation as main SKU
                            if _should_rotate_sku(stack_sku, rot):
                                stack_sku.width, stack_sku.height = stack_sku.height, stack_sku.width
                            bundle.add_sku(stack_sku, x, y, stack_rotated)
                        
                        # Remove all placed SKUs from remaining list
                        remaining_skus.remove(sku)
                        for stack_sku in stackable_skus:
                            if stack_sku in remaining_skus:
                                remaining_skus.remove(stack_sku)
                        
                        placed = True
                        break
                if placed:
                    break
            if placed:
                break
    return remaining_skus

def find_stackable_skus(target_sku: SKU, remaining_skus: List[SKU], unavailable_skus: set, target_index: int, max_length: int) -> List[SKU]:
    """Find SKUs that can be stacked with target SKU based on combined length"""
    stackable = []
    current_total_length = target_sku.length
    
    # Create a list of candidates with their indices for sorting
    candidates = []
    for i, sku in enumerate(remaining_skus):
        if id(sku) in unavailable_skus:
            continue
        if i != target_index and _skus_compatible_for_stacking(sku, target_sku):
            candidates.append((sku, i))
    
    # Sort candidates by length (descending) to place larger SKUs in back
    candidates.sort(key=lambda x: x[0].length, reverse=True)
    
    # Add SKUs to stack as long as they fit within max_length
    for sku, _ in candidates:
        if current_total_length + sku.length <= max_length:
            stackable.append(sku)
            current_total_length += sku.length
        # Stop if we've reached the maximum length
        if current_total_length >= max_length:
            break
    
    return stackable

def _skus_compatible_for_stacking(sku1: SKU, sku2: SKU) -> bool:
    """Check if two SKUs are compatible for stacking based on dimensions and properties"""
    # Check if candidate sku is within 13mm of target sku width and height
    return (abs(sku1.width - sku2.width) <= 13 and
            abs(sku1.height - sku2.height) <= 13)

def _sku_within_height_range(sku: SKU, row_skus: List) -> bool:
    """Check if SKU is within height tolerance of previous SKU"""
    height_tol = 100 # mm tolerance for height matching
    if not row_skus:
        return True
    last_sku_height = row_skus[-1][1].height
    return (last_sku_height - height_tol <= sku.height <= last_sku_height + height_tol)