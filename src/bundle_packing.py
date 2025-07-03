from typing import List, Tuple
from bundle_classes import SKU, Bundle, PlacedSKU, FILLER_62, FILLER_44

global maxLength
maxLength = 3680  # Maximum length for bundles

def pack_skus_with_pattern(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """
    Packs SKUs into bundles using a pattern-based algorithm with length stacking support
    """

    def has_sufficient_support(x, y, width, bundle: Bundle, threshold=0.7) -> bool:
        """
        Check if at least `threshold` fraction of the bottom width is supported by other SKUs below.
        """
        support_segments = []
        for sku in bundle.skus:
            if sku.y + sku.height == y:  # SKU is directly underneath
                overlap_x_start = max(x, sku.x)
                overlap_x_end = min(x + width, sku.x + sku.width)
                if overlap_x_end > overlap_x_start:
                    support_segments.append((overlap_x_start, overlap_x_end))

        total_supported_width = sum(end - start for start, end in support_segments)
        return (total_supported_width / width) >= threshold

    def get_sku_dimensions(sku: SKU, vertical: bool):
        """Return (width, height) for SKU based on orientation"""
        if vertical:
            # Vertical: shortest side as width, longest as height
            return min(sku.width, sku.height), max(sku.width, sku.height)
        else:
            # Horizontal: longest side as width, shortest as height
            return max(sku.width, sku.height), min(sku.width, sku.height)

    def can_fit_in_bundle(sku: SKU, x: int, y: int, vertical: bool, bundle: Bundle) -> bool:
        """Check if SKU can fit at position with given orientation"""
        width, height = get_sku_dimensions(sku, vertical)

        # Check bundle boundaries
        if x + width > bundle.width or y + height > bundle.height:
            return False

        # Check for overlap with existing SKUs
        for placed_sku in bundle.skus:
            if (x < placed_sku.x + placed_sku.width and
                x + width > placed_sku.x and
                y < placed_sku.y + placed_sku.height and
                y + height > placed_sku.y):
                return False
        return True

    def can_fit_any_orientation(sku: SKU, x: int, y: int, bundle: Bundle) -> tuple:
        """Check if SKU can fit at position in any orientation. Returns (can_fit, rotated)"""
        # Try original orientation
        if (x + sku.width <= bundle.width and y + sku.height <= bundle.height):
            overlap = False
            for placed_sku in bundle.skus:
                if (x < placed_sku.x + placed_sku.width and
                    x + sku.width > placed_sku.x and
                    y < placed_sku.y + placed_sku.height and
                    y + sku.height > placed_sku.y):
                    overlap = True
                    break
            if not overlap:
                return True, False

        # Try rotated orientation
        if (x + sku.height <= bundle.width and y + sku.width <= bundle.height):
            overlap = False
            for placed_sku in bundle.skus:
                if (x < placed_sku.x + placed_sku.width and
                    x + sku.height > placed_sku.x and
                    y < placed_sku.y + placed_sku.height and
                    y + sku.width > placed_sku.y):
                    overlap = True
                    break
            if not overlap:
                return True, True

        return False, False

    def find_stackable_skus(target_sku: SKU, remaining_skus: List[SKU], max_quantity: int) -> List[SKU]:
        """Find SKUs that can be stacked with the target SKU based on length constraints"""
        stackable = []
        max_stack = calculate_max_stack_quantity(target_sku.length, maxLength)
        
        if max_stack <= 1:
            return stackable
            
        # Look for identical SKUs to stack
        count = 0
        for sku in remaining_skus:
            if (sku.id == target_sku.id and 
                sku.width == target_sku.width and 
                sku.height == target_sku.height and
                sku.length == target_sku.length and
                count < min(max_stack - 1, max_quantity - 1)):  # -1 because we already have the target
                stackable.append(sku)
                count += 1
                
        return stackable[:-1]

    def greedy_fill_remaining_space(bundle: Bundle, remaining_skus: List[SKU]) -> List[SKU]:
        """Fill remaining space in bundle using greedy algorithm with stacking"""
        still_remaining = remaining_skus.copy()

        # Generate candidate points from existing SKUs
        candidate_points = {(0, 0)}
        for placed_sku in bundle.skus:
            candidate_points.add((placed_sku.x + placed_sku.width, placed_sku.y))
            candidate_points.add((placed_sku.x, placed_sku.y + placed_sku.height))

        # Sort candidate points by y first, then x (bottom-left preference)
        sorted_points = sorted(candidate_points, key=lambda p: (p[1], p[0]))

        placed_any = True
        while placed_any and still_remaining:
            placed_any = False

            for point in sorted_points:
                x, y = point
                best_sku_idx = None
                best_rotation = False
                best_stack_quantity = 1

                # Find the best SKU for this position (prefer heavier SKUs)
                for i, sku in enumerate(still_remaining):
                    can_fit, rotated = can_fit_any_orientation(sku, x, y, bundle)
                    if can_fit:
                        # Calculate how many we can stack
                        stackable_skus = find_stackable_skus(sku, still_remaining, len(still_remaining))
                        stack_quantity = len(stackable_skus) + 1
                        
                        if best_sku_idx is None or sku.weight * stack_quantity > still_remaining[best_sku_idx].weight * best_stack_quantity:
                            best_sku_idx = i
                            best_rotation = rotated
                            best_stack_quantity = stack_quantity

                if best_sku_idx is not None:
                    # Place the best SKU with stacking
                    sku = still_remaining.pop(best_sku_idx)
                    
                    # Find stackable SKUs and their indices
                    stackable_indices = []
                    stackable_count = 0
                    for i, remaining_sku in enumerate(still_remaining):
                        if (remaining_sku.id == sku.id and 
                            remaining_sku.width == sku.width and 
                            remaining_sku.height == sku.height and
                            remaining_sku.length == sku.length and
                            stackable_count < best_stack_quantity - 1):
                            stackable_indices.append(i)
                            stackable_count += 1

                    # Remove stackable SKUs by index (in reverse order to maintain indices)
                    for idx in sorted(stackable_indices, reverse=True):
                        still_remaining.pop(idx)

                    if best_rotation:
                        sku.width, sku.height = sku.height, sku.width

                    bundle.add_sku(sku, x, y, best_rotation, best_stack_quantity)

                    # Update candidate points
                    candidate_points.add((x + sku.width, y))
                    candidate_points.add((x, y + sku.height))
                    sorted_points = sorted(candidate_points, key=lambda p: (p[1], p[0]))

                    placed_any = True
                    break

        return still_remaining

    def pack_bundle_with_pattern(available_skus: List[SKU], bundle: Bundle) -> List[SKU]:
        """
        Pack a single bundle following the vertical/horizontal pattern with stacking
        """
        remaining_skus = available_skus.copy()
        current_y = 0
        is_vertical_row = True
        vertical_row_height = 0
        horizontal_start_y = 0

        while remaining_skus and current_y < bundle.height:
            # Find SKUs that can fit in current row orientation
            row_skus = []
            row_height = 0
            current_x = 0

            # Sort SKUs by weight (heaviest first) for each row
            remaining_skus.sort(key=lambda x: x.weight, reverse=True)

            # Track which SKUs we've already considered for this row to avoid duplicates
            considered_skus = set()

            for i, sku in enumerate(remaining_skus):
                # Skip if we've already considered this SKU for stacking
                if id(sku) in considered_skus:
                    continue
                    
                width, height = get_sku_dimensions(sku, is_vertical_row)

                # Check if SKU fits in current row
                if (current_x + width <= bundle.width and
                    current_y + height <= bundle.height and
                    can_fit_in_bundle(sku, current_x, current_y, is_vertical_row, bundle)):
                    
                    # Simulate placement to verify support
                    x = current_x
                    y = current_y
                    if current_y != 0:  # not first row
                        if not has_sufficient_support(x, y, width, bundle):
                            break  # Not enough support; start a new row

                    # Calculate stacking - find stackable SKUs from remaining list
                    stackable_skus = []
                    max_stack = calculate_max_stack_quantity(sku.length, maxLength)
                    
                    if max_stack > 1:
                        # Look for identical SKUs to stack (excluding the current one)
                        for j, other_sku in enumerate(remaining_skus):
                            if (j != i and  # Don't include the current SKU
                                other_sku.id == sku.id and 
                                other_sku.width == sku.width and 
                                other_sku.height == sku.height and
                                other_sku.length == sku.length and
                                len(stackable_skus) < max_stack - 1):  # -1 because we already have the base SKU
                                stackable_skus.append(other_sku)
                    
                    stack_quantity = len(stackable_skus) + 1

                    # Mark all SKUs in this stack as considered
                    considered_skus.add(id(sku))
                    for stackable_sku in stackable_skus:
                        considered_skus.add(id(stackable_sku))

                    row_skus.append((i, sku, current_x, current_y, width, height, stack_quantity, stackable_skus))
                    current_x += width
                    row_height = max(row_height, height)

            if not row_skus:
                break  # No more SKUs fit

            # Place the SKUs in this row and remove them from remaining_skus
            skus_to_remove = []
            
            for i, sku, x, y, width, height, stack_quantity, stackable_skus in row_skus:
                # Determine if rotation is needed
                original_width, original_height = sku.width, sku.height
                rotated = False

                if is_vertical_row:
                    # Vertical row: want min dimension as width
                    if original_width > original_height:
                        rotated = True
                        sku.width, sku.height = sku.height, sku.width
                else:
                    # Horizontal row: want max dimension as width
                    if original_width < original_height:
                        rotated = True
                        sku.width, sku.height = sku.height, sku.width

                bundle.add_sku(sku, x, y, rotated, stack_quantity)
                
                # Collect SKUs to remove (base SKU + stackable SKUs)
                skus_to_remove.append(sku)
                skus_to_remove.extend(stackable_skus)

            # Remove all placed SKUs from remaining_skus
            for sku_to_remove in skus_to_remove:
                if sku_to_remove in remaining_skus:
                    remaining_skus.remove(sku_to_remove)

            # Move to next row
            current_y += row_height

            # Determine orientation for next row
            if is_vertical_row:
                # After placing a vertical row, switch to horizontal and record the height
                is_vertical_row = False
                vertical_row_height = row_height
                horizontal_start_y = current_y  # Start tracking horizontal section from current position
            else:
                # We're in horizontal mode - check if we should switch back to vertical
                # Calculate total height of horizontal section so far
                horizontal_section_height = current_y - horizontal_start_y

                # Only switch back to vertical when horizontal section matches vertical row height
                if horizontal_section_height >= vertical_row_height:
                    is_vertical_row = True
                    # Reset for next cycle
                    vertical_row_height = 0
                    horizontal_start_y = 0

        return remaining_skus

    # Main packing logic
    skus.sort(key=lambda x: x.weight, reverse=True)  # Sort by weight initially
    bundles: List[Bundle] = []
    remaining_skus = skus.copy()

    while remaining_skus:
        new_bundle = Bundle(bundle_width, bundle_height, maxLength)

        # Check if any remaining SKU can fit in a bundle at all
        can_fit_any = False
        for sku in remaining_skus:
            vert_w, vert_h = get_sku_dimensions(sku, True)
            horiz_w, horiz_h = get_sku_dimensions(sku, False)

            if ((vert_w <= bundle_width and vert_h <= bundle_height) or
                (horiz_w <= bundle_width and horiz_h <= bundle_height)):
                can_fit_any = True
                break

        if not can_fit_any:
            break

        # Pack this bundle with pattern-based algorithm
        before_count = len(remaining_skus)
        remaining_skus = pack_bundle_with_pattern(remaining_skus, new_bundle)
        after_count = len(remaining_skus)

        # GREEDY OPTIMIZATION: Try to fill remaining space with leftover SKUs
        if remaining_skus and new_bundle.skus:
            remaining_skus = greedy_fill_remaining_space(new_bundle, remaining_skus)

        if before_count == after_count:
            # No progress made, skip the largest remaining SKU
            if remaining_skus:
                largest_sku = max(remaining_skus, key=lambda x: x.width * x.height)
                remaining_skus.remove(largest_sku)

        if new_bundle.skus:  # Only add bundles that have SKUs
            new_bundle.add_packaging() # Add packaging to bundle (increases weight)
            bundles.append(new_bundle)

    return bundles

def add_filler_material(bundle: Bundle) -> None:
    """
    Add filler material to empty spaces in the bundle - aggressive filling
    """
    if not bundle.skus:
        return
    
    fillers = [FILLER_62, FILLER_44]  # Try 62 first as it's more efficient space-wise
    
    # Keep trying to place filler until no more can be placed
    placed_any = True
    while placed_any:
        placed_any = False
        
        # Generate all possible candidate points more comprehensively
        candidate_points = set()
        
        # Add origin
        candidate_points.add((0, 0))
        
        # Add points at all corners and edges of existing SKUs
        for placed_sku in bundle.skus:
            # Right edge of SKU
            candidate_points.add((placed_sku.x + placed_sku.width, placed_sku.y))
            candidate_points.add((placed_sku.x + placed_sku.width, placed_sku.y + placed_sku.height))
            
            # Top edge of SKU
            candidate_points.add((placed_sku.x, placed_sku.y + placed_sku.height))
            candidate_points.add((placed_sku.x + placed_sku.width, placed_sku.y + placed_sku.height))
            
            # Also add points slightly inside for better coverage
            if placed_sku.x > 0:
                candidate_points.add((placed_sku.x, placed_sku.y))
            if placed_sku.y > 0:
                candidate_points.add((placed_sku.x, placed_sku.y))

        # Add a grid of points for comprehensive coverage
        grid_size = 25  # Check every 25mm
        for x in range(0, int(bundle.width), grid_size):
            for y in range(0, int(bundle.height), grid_size):
                candidate_points.add((x, y))

        # Sort candidate points by area potential (prefer larger empty areas)
        # Calculate potential area for each point
        point_priorities = []
        for point in candidate_points:
            x, y = point
            # Calculate maximum possible rectangle from this point
            max_width = bundle.width - x
            max_height = bundle.height - y
            
            # Check how much space is actually available
            for placed_sku in bundle.skus:
                if placed_sku.x >= x and placed_sku.y >= y:
                    # SKU is to the right and/or above this point
                    if placed_sku.x < x + max_width:
                        max_width = min(max_width, placed_sku.x - x)
                    if placed_sku.y < y + max_height:
                        max_height = min(max_height, placed_sku.y - y)
            
            potential_area = max_width * max_height
            point_priorities.append((potential_area, x, y))
        
        # Sort by potential area (largest first), then by position
        point_priorities.sort(key=lambda p: (-p[0], p[2], p[1]))
        
        for _, x, y in point_priorities:
            best_filler = None
            best_config = None
            best_area = 0
            
            for filler in fillers:
                # Try both orientations for filler
                orientations = [
                    (filler.width, filler.height, False),
                    (filler.height, filler.width, True)
                ]
                
                for width, height, rotated in orientations:
                    # Check if filler fits
                    if (x + width <= bundle.width and y + height <= bundle.height):
                        # Check for overlap
                        overlap = False
                        for placed_sku in bundle.skus:
                            if (x < placed_sku.x + placed_sku.width and
                                x + width > placed_sku.x and
                                y < placed_sku.y + placed_sku.height and
                                y + height > placed_sku.y):
                                overlap = True
                                break
                        
                        if not overlap:
                            area = width * height
                            if area > best_area:
                                best_area = area
                                best_filler = filler
                                best_config = (width, height, rotated)
            
            if best_filler and best_config:
                width, height, rotated = best_config
                
                # Calculate how many fillers can be stacked lengthwise
                stack_quantity = calculate_max_stack_quantity(best_filler.length, maxLength)
                
                # Create a copy of filler with potentially rotated dimensions
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
                break  # Start over with updated bundle state

def calculate_max_stack_quantity(sku_length: float, max_bundle_length: float) -> int:
    """
    Calculate how many SKUs can be stacked in the length dimension
    """
    return int(max_bundle_length // sku_length)

def pack_skus(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    """
    Wrapper function that packs SKUs and adds filler material
    """
    bundles = pack_skus_with_pattern(skus, bundle_width, bundle_height)

    # Add filler material and resize bundles
    for bundle in bundles:
        bundle.resize_to_content()
        add_filler_material(bundle)

    return bundles
