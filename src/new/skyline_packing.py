from typing import List, Tuple
from bundle_classes import Bundle, SKU

def fill_remaining_space_with_skyline(bundle: Bundle, remaining_skus: List[SKU]) -> List[SKU]:
    """
    Fill remaining space using skyline algorithm - ENHANCED VERSION
    """
    # Sort SKUs by largest dimension (or weight for better packing)
    remaining_skus.sort(key=lambda s: (max(s.height, s.width), s.weight), reverse=True)
    
    placed_any = True
    
    while placed_any and remaining_skus:
        placed_any = False
        
        # Get current bundle surface profile
        profile = get_top_surface_profile(bundle, bundle.width)
        
        if not profile:
            # If no profile (empty bundle), try placing at origin
            candidate_points = [(0, 0)]
        else:
            # Find candidate placement points from the profile
            candidate_points = []
            
            # Add all profile points as candidates
            for x, y in profile:
                candidate_points.append((x, y))
            
            # Also add points where placed SKUs end
            for sku in bundle.skus:
                candidate_points.append((sku.x + sku.width, sku.y))
                candidate_points.append((sku.x, sku.y + sku.height))
            
            # Remove duplicates and sort by y-coordinate, then x-coordinate
            candidate_points = list(set(candidate_points))
            candidate_points.sort(key=lambda p: (p[1], p[0]))
        
        # Try to place each remaining SKU at each candidate point
        for x, y in candidate_points:
            if placed_any:
                break
                
            for i, sku in enumerate(remaining_skus):
                if placed_any:
                    break
                
                # Check bottom row restrictions
                if y == 0:
                    if not sku.can_be_bottom:
                        continue
                    # Check length requirements for bottom row
                    if bundle.max_length == 7340 and abs(sku.length - 7340) > 100:
                        continue
                    if bundle.max_length == 3680 and abs(sku.length - 3680) > 100:
                        continue
                
                # Determine best orientation to minimize height impact
                best_placement = _find_best_orientation_for_placement(sku, x, y, bundle, remaining_skus, i)
                
                if not best_placement:
                    continue
                
                width, height, rotated, stack_quantity, stackable_skus = best_placement
                
                # Try to slide the SKU left for better packing
                optimized_x = find_snug_x_left(x, y, width, height, bundle, rotated)
                
                # Place the SKU
                if rotated:
                    sku.width, sku.height = sku.height, sku.width
                
                bundle.add_sku(sku, optimized_x, y, rotated, stack_quantity)
                
                # Remove the placed SKU and its stackable companions
                remaining_skus.remove(sku)
                for stackable_sku in stackable_skus:
                    if stackable_sku in remaining_skus:
                        remaining_skus.remove(stackable_sku)
                
                placed_any = True
                break
                    
    return remaining_skus

# SKYLINE HELPER FUNCTIONS
def _find_best_orientation_for_placement(sku: SKU, x: int, y: int, bundle: Bundle, 
                                       remaining_skus: List[SKU], sku_index: int) -> tuple:
    """
    Find the best orientation for placing a SKU to minimize height impact and white space
    """
    orientations = [
        (sku.width, sku.height, False),   # Original orientation
        (sku.height, sku.width, True)    # Rotated orientation
    ]
    
    best_placement = None
    best_score = float('inf')
    
    for width, height, rotated in orientations:
        # Check if SKU fits within bundle boundaries
        if x + width > bundle.width or y + height > bundle.height:
            continue
        
        # Check for overlaps with existing SKUs
        if _has_overlap_with_existing_skus(x, y, width, height, bundle):
            continue
        
        # Check if there's sufficient support (except for bottom row)
        if y > 0 and not has_sufficient_support(x, y, width, bundle):
            continue
        
        # Find stackable SKUs for this position
        stackable_skus = find_stackable_skus(sku, remaining_skus, sku_index, bundle.max_length)
        stack_quantity = len(stackable_skus) + 1
        
        # Calculate placement score (lower is better)
        score = _calculate_placement_score(x, y, width, height, bundle)
        
        if score < best_score:
            best_score = score
            best_placement = (width, height, rotated, stack_quantity, stackable_skus)
    
    return best_placement

def _calculate_placement_score(x: int, y: int, width: int, height: int, bundle: Bundle) -> float:
    """
    Calculate a score for SKU placement (lower is better)
    Prioritizes: lower height, less white space creation, better space utilization
    """
    # Height penalty (placing higher is worse)
    height_penalty = (y + height) * 10
    
    # White space penalty - calculate potential wasted space above this placement
    white_space_penalty = _calculate_white_space_penalty(x, y, width, height, bundle)
    
    # Edge preference (prefer placing against existing SKUs)
    edge_bonus = 0
    if x == 0:  # Left edge
        edge_bonus -= 5
    
    # Check if this placement creates a good "shelf" for future placements
    shelf_bonus = 0
    if width >= 50:  # Good shelf width
        shelf_bonus -= 3
    
    return height_penalty + white_space_penalty + edge_bonus + shelf_bonus

def _calculate_white_space_penalty(x: int, y: int, width: int, height: int, bundle: Bundle) -> float:
    """
    Calculate penalty for potential white space creation
    """
    penalty = 0
    
    # Check if this placement creates unusable thin gaps
    for sku in bundle.skus:
        # Check for thin vertical gaps
        if sku.x + sku.width == x and abs(sku.y - y) < height:
            gap_width = x - (sku.x + sku.width)
            if 0 < gap_width < 30:  # Unusable thin gap
                penalty += 20
        
        # Check for thin horizontal gaps
        if sku.y + sku.height == y and abs(sku.x - x) < width:
            gap_height = y - (sku.y + sku.height)
            if 0 < gap_height < 30:  # Unusable thin gap
                penalty += 20
    
    return penalty

def find_snug_x_left(x: float, y: float, w: float, h: float, bundle: Bundle, rotated: bool) -> float:
    """
    Slide the SKU left until it either hits the left wall or another SKU, or loses sufficient support.
    """
    original_support = has_sufficient_support(x, y, w, bundle, ret_real_val=True)
    
    while x > 0:
        next_x = x - 1
        
        # Check if support is still sufficient and not degrading
        if not has_sufficient_support(next_x, y, w, bundle):
            break
        
        current_support = has_sufficient_support(next_x, y, w, bundle, ret_real_val=True)
        if current_support < original_support * 0.9:  # Allow slight degradation
            break
        
        # Check for overlap with existing SKUs
        if _has_overlap_with_existing_skus(next_x, y, w, h, bundle):
            break
        
        x = next_x
    
    return x

def _has_overlap_with_existing_skus(x: int, y: int, width: int, height: int, bundle: Bundle) -> bool:
    """
    Check if a rectangle at (x, y) with given width and height overlaps with any existing SKUs
    """
    for placed_sku in bundle.skus:
        # Check if rectangles overlap
        if (x < placed_sku.x + placed_sku.width and
            x + width > placed_sku.x and
            y < placed_sku.y + placed_sku.height and
            y + height > placed_sku.y):
            return True
    return False

def get_top_surface_profile(bundle: Bundle, max_width: float) -> List[Tuple[float, float]]:
    """
    Returns top surface profile at key points only (more efficient)
    """
    if not bundle.skus:
        return [(0, 0)]
    
    # Collect all x-coordinates where height might change
    x_coords = set([0])
    for sku in bundle.skus:
        x_coords.add(sku.x)
        x_coords.add(sku.x + sku.width)
    
    # Add max_width to ensure we cover the full width
    x_coords.add(max_width)
    
    # Sort x-coordinates
    x_coords = sorted(x_coords)
    
    surface = []
    for x in x_coords:
        if x >= max_width:
            break
            
        max_y = 0
        for sku in bundle.skus:
            if sku.x <= x < sku.x + sku.width:
                max_y = max(max_y, sku.y + sku.height)
        surface.append((x, max_y))
    
    return surface

def find_lowest_points(profile: List[Tuple[float, float]]) -> List[Tuple[float, float]]:
    """
    Returns points sorted by lowest height, then by x-coordinate
    """
    if not profile:
        return [(0, 0)]
    
    return sorted(profile, key=lambda p: (p[1], p[0]))

def has_sufficient_support(x: float, y: float, width: float, bundle: Bundle, 
                          threshold: float = 0.7, max_gap: float = 10, ret_real_val: bool = False) -> bool:
    """
    Checks if at least threshold fraction of width is supported
    Enhanced version with better gap handling and optional real value return
    """
    if y == 0:  # Bottom row always has support
        return True if not ret_real_val else 1.0
        
    support_segments = []
    
    # Find all SKUs that could provide support
    for sku in bundle.skus:
        sku_top = sku.y + sku.height
        
        # Check if this SKU is close enough to provide support
        if abs(sku_top - y) <= max_gap:
            overlap_start = max(x, sku.x)
            overlap_end = min(x + width, sku.x + sku.width)
            if overlap_end > overlap_start:
                support_segments.append((overlap_start, overlap_end))
    
    if not support_segments:
        return False if not ret_real_val else 0.0
    
    # Merge overlapping segments
    support_segments.sort()
    merged = []
    for start, end in support_segments:
        if not merged or start > merged[-1][1]:
            merged.append((start, end))
        else:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
    
    total_supported = sum(end - start for start, end in merged)
    support_ratio = total_supported / width
    
    if ret_real_val:
        return support_ratio
    else:
        return support_ratio >= threshold

def find_stackable_skus(target_sku: SKU, remaining_skus: List[SKU], target_index: int, max_length: int) -> List[SKU]:
    """Find SKUs that can be stacked with target SKU"""
    stackable = []
    max_stack = _calculate_max_stack_quantity(target_sku.length, max_length)

    if max_stack <= 1:
        return stackable

    for i, sku in enumerate(remaining_skus):
        if (i != target_index and _skus_identical(sku, target_sku) and 
            len(stackable) < max_stack - 1):
            stackable.append(sku)

    return stackable

def _skus_identical(sku1: SKU, sku2: SKU) -> bool:
    """Check if two SKUs are identical for stacking purposes"""
    return (sku1.id == sku2.id and 
            sku1.width == sku2.width and 
            sku1.height == sku2.height and
            sku1.length == sku2.length)

def _calculate_max_stack_quantity(sku_length: float, max_bundle_length: float) -> int:
    """Calculate maximum stack quantity based on length"""
    return int(max_bundle_length // sku_length)