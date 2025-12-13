import copy
from typing import List, Tuple
from bundle_classes import SKU, Bundle, FILLER_44, FILLER_62
from getJSONdata import VARIABLES
from collections import Counter

MAX_LENGTH = 3680 # mm
MAX_WEIGHT = round(VARIABLES['MAX_WEIGHT'])  # kg
MIN_HEIGHT_WIDTH_RATIO = VARIABLES['MIN_HEIGHT_WIDTH_RATIO']  # Minimum height-to-width ratio for bundles
BOTTOM_ROW_LENGTH = 0
REMOVED_SKUS = []

MIN_CEILING_COVERAGE = VARIABLES['MIN_CEILING_COVERAGE']  # Minimum ceiling coverage required (%)
MAX_DIST_FROM_CEILING = round(VARIABLES['MAX_DIST_FROM_CEILING'])  # mm, maximum distance from ceiling to be considered sufficient coverage
STACKING_MAX_DIFF = round(VARIABLES['STACKING_MAX_DIFF'])  # mm, maximum difference in width and height for lengthwise stacking SKUs
SKU_MAX_HEIGHT_DIFF = round(VARIABLES['SKU_MAX_HEIGHT_DIFF'])  # mm, maximum height difference for SKUs to be considered compatible in a row
BASE_COVERAGE_THRESHOLD = VARIABLES['BASE_COVERAGE_THRESHOLD']  # Base coverage threshold for support checks (%)
SKU_COVERAGE_HEIGHT_BUFFER = round(VARIABLES['SKU_COVERAGE_HEIGHT_BUFFER'])  # mm, buffer for how far below a SKU can be to be considered coverage

def pack_skus(skus: List[SKU], bundle_width: int, bundle_height: int, mach1_skus: List[str]) -> List[Bundle]:
    """Main entry point for packing SKUs into bundles"""
    global FILLER_62, FILLER_44
    from bundle_classes import FILLER_44, FILLER_62
    # Separate SKUs with bundle override
    override_skus = [sku for sku in skus if sku.data and sku.data.get('Bdl_Override')]
    component_skus = [sku for sku in skus if sku.data and sku.data.get('Component') and not sku.data.get('Bdl_Override')]
    regular_skus = [sku for sku in skus if (sku not in override_skus and sku not in component_skus)]

    # Process override bundles first
    override_bundles = _process_override_bundles(override_skus, bundle_width, bundle_height, mach1_skus)
    if override_bundles == -1:
        return -1, REMOVED_SKUS

    # Pack component SKUs into their own bundles
    # Find the machine that the SKUs belong to; if both, use "MIXED"
    mach1_count = [sku[-3:] in mach1_skus for sku in [s.id for s in component_skus]].count(True)
    if mach1_count == 0:
        machine = 'MACH5'
    elif mach1_count == len(component_skus):
        machine = 'MACH1'
    else:
        machine = 'MIXED'
    component_bundles = _pack_skus_with_pattern(component_skus, bundle_width, bundle_height, machine=machine)

    # Group SKUs by color
    color_groups = _group_skus_by_color(regular_skus)
    remaining_color_groups = list(color_groups.items())

    can_try_merge_bundles_mach5 = []
    can_try_merge_bundles_mach1 = []
    final_bundles = []

    while remaining_color_groups:
        color, color_skus = remaining_color_groups.pop(0)
        color_skus.sort(key=lambda x: max(x.width, x.height), reverse=True)

        if color[-3:] in mach1_skus:
            base_bundles = _pack_skus_with_pattern(color_skus, bundle_width, bundle_height, machine='MACH1')
            can_try_merge_bundles_mach1.extend(base_bundles)
        else:
            base_bundles = _pack_skus_with_pattern(color_skus, bundle_width, bundle_height, machine='MACH5')
            can_try_merge_bundles_mach5.extend(base_bundles)

    # try to merge bundles if they can all fit in one
    merged_bundles_mach1 = _try_merge_bundles(can_try_merge_bundles_mach1, bundle_width, bundle_height, machine='MACH1')
    merged_bundles_mach5 = _try_merge_bundles(can_try_merge_bundles_mach5, bundle_width, bundle_height, machine='MACH5')
    merged_machine_bundles = merged_bundles_mach1 + merged_bundles_mach5

    final_bundles = _try_merge_bundles(merged_machine_bundles, bundle_width, bundle_height, machine='MACH5', diff_machines=True) # mach5 as placeholder
    # for bundle in component_bundles:
    #     if bundle.packing_machine == 'MACH1':
    #         # can_try_merge_bundles_mach1.append(bundle)
    #         mach1_component_bundles = _fill_bundle_with_components(bundle, final_bundles)
    #     else:
    #         # can_try_merge_bundles_mach5.append(bundle)
    #         mach5_component_bundles = _fill_bundle_with_components(bundle, final_bundles)
    extra_component_bundles = _fill_bundles_with_components(component_bundles, final_bundles)
    final_bundles.extend(_try_merge_bundles(extra_component_bundles, bundle_width, bundle_height, machine=machine))

    # remove any empty bundles
    final_bundles = [bundle for bundle in final_bundles if (bundle.width > 0 and bundle.height > 0)]
    for bundle in final_bundles:
        bundle.add_packaging()  # Add packaging to each bundle
    for bundle in override_bundles:
        bundle.add_packaging()

    return override_bundles + final_bundles, REMOVED_SKUS

def _fill_bundles_with_components(component_bundles: List[Bundle], target_bundles: List[Bundle]) -> List[Bundle]:
    """Place as many component SKUs on top of other SKUs in target bundles as possible"""


def _try_merge_bundles(bundles: List[Bundle], bundle_width: int, bundle_height: int, machine: str, diff_machines: bool = False) -> List[Bundle]:
    """Attempt to merge bundles if they can all fit in one bundle"""
    attempted_merged_bundles = []
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

                # identify if bundles have already been attempted to be merged
                if ([id(bundle1), id(bundle2)] in attempted_merged_bundles
                    or [id(bundle2), id(bundle1)] in attempted_merged_bundles
                    # flag to merge across machines is true, don't check when machine is the same (since they have already been done)
                    or (diff_machines and bundle1.packing_machine == bundle2.packing_machine)):
                    continue
                else:
                    attempted_merged_bundles.append([id(bundle1), id(bundle2)])

                # get all SKUs from both bundles
                all_skus = bundle1.skus + bundle2.skus
                # deconstruct PlacedSKU objects into just SKU objects for consistency
                all_skus = [sku for sku in all_skus if "Filler" not in sku.id]

                # preliminary check if they can fit into one bundle
                bundle1_area = bundle1.width * bundle1.height
                bundle2_area = bundle2.width * bundle2.height
                if (bundle1_area + bundle2_area > bundle_width * bundle_height
                    or (bundle1.get_total_weight() + bundle2.get_total_weight() > MAX_WEIGHT)):
                    continue
                # try to pack them into a new bundle
                merged_bundles = _pack_skus_with_pattern(all_skus, bundle_width, bundle_height, machine=machine, merging=True)
                if len(merged_bundles) == 1:
                    # if they fit into one bundle, remove the original bundles
                    bundles.pop(j)
                    bundles.pop(i)
                    if diff_machines:
                        merged_bundles[0].packing_machine = 'MIXED'
                    bundles.append(merged_bundles[0])
                    merging_able = True
                    break
            if merging_able:
                break
    # After merging, create combined bundles if they are laid flat
    sku_groups = {}
    flat_count = []
    for bundle_idx, bundle in enumerate(reversed(bundles)):
        if all(not sku.rotated for sku in bundle.skus):
            flat_count.append(bundle)
            # skus are laid flat
            for sku in bundle.skus:
                if f"{bundle_idx}_{sku.y}" not in sku_groups:
                    sku_groups[f"{bundle_idx}_{sku.y}"] = []
                sku_groups[f"{bundle_idx}_{sku.y}"].append(sku)
            bundles.remove(bundle)
        else:
            _add_filler_material(bundle)
    if len(flat_count) == 1:
        _add_filler_material(flat_count[0])
        bundles.append(flat_count[0])
        sku_groups = {}

    flat_bundle = Bundle(bundle_width, bundle_height, MAX_LENGTH, packing_machine=machine)
    if sku_groups:
        for group_key, group_skus in sku_groups.items():
            flat_bundle.skus.extend(group_skus)
        bundles.extend(_stack_skus_flat(flat_bundle, sku_groups))

    return bundles

def _pack_skus_with_pattern(skus: List[SKU], bundle_width: int, bundle_height: int, merging: bool = False, machine: str = 'MACH5') -> List[Bundle]:
    """Pack SKUs into bundles using pattern-based algorithm"""
    global REMOVED_SKUS
    if not skus:
        return []
    skus.sort(key=lambda x: max(x.height, x.width), reverse=True)

    bundles = []
    remaining_skus = skus.copy()

    while remaining_skus:
        if merging and len(bundles) > 0:
            return [bundles[0], bundles[0]] # 1 bundle and remaining skus, can't be merged so stop trying
        max_length = 3680 if max(sku.length for sku in remaining_skus if sku.length) < 3700 else 7340
        # Check if any SKU can fit in a bundle
        if not any((sku.can_be_bottom and abs(sku.length - max_length) <= 100) for sku in remaining_skus):
            # If no SKU can be bottom, set all to True
            for sku in remaining_skus:
                sku.can_be_bottom = True
        if not _can_any_sku_fit(remaining_skus, bundle_width, bundle_height):
            break

        temp_width = bundle_width
        temp_height = bundle_height
        before_count = len(remaining_skus)
        skus_copy = copy.deepcopy(remaining_skus)
        new_bundle = False

        while True: # Create first iteration
            if new_bundle:
                bundle = new_bundle
                new_bundle = False
            else:
                bundle = Bundle(temp_width, temp_height, MAX_LENGTH, machine)
                remaining_skus = _pack_single_bundle(skus_copy, bundle)

            # If height is 0.3x width or lower, reduce width and try again
            if (bundle.height / bundle.width < MIN_HEIGHT_WIDTH_RATIO and len(bundle.skus) > 2):
                temp_width = round(bundle.width - 20)
                continue
            if (not _has_sufficient_ceiling_coverage(bundle) or bundle.height > bundle.width):
                any_sku_not_bottom = False
                for sku in reversed(bundle.skus):
                    if sku.y != 0:
                        any_sku_not_bottom = True
                        break
                if any_sku_not_bottom:
                    # try repacking with both reduced height and reduced width, see which one gets better coverage
                    # reduce height
                    max_sku = max(bundle.skus, key=lambda s: s.y + s.height, default=None)
                    temp_temp_height = round(bundle.height - min(max_sku.height + 1, 20))
                    bundle_reduced_height = Bundle(temp_width, temp_temp_height, MAX_LENGTH, packing_machine=machine)
                    rs1 = _pack_single_bundle(skus_copy, bundle_reduced_height)
                    height_ceiling_coverage = _has_sufficient_ceiling_coverage(bundle_reduced_height, get_value=True)

                    # reduce width
                    max_sku = max(bundle.skus, key=lambda s: s.x + s.width, default=None)
                    temp_temp_width = round(bundle.width - min(max_sku.width + 1, 20))
                    bundle_reduced_width = Bundle(temp_temp_width, temp_height, MAX_LENGTH, packing_machine=machine)
                    rs2 = _pack_single_bundle(skus_copy, bundle_reduced_width)
                    width_ceiling_coverage = _has_sufficient_ceiling_coverage(bundle_reduced_width, get_value=True)

                    # compare (if one has more skus packed, pick that one; if same, pick one with better ceiling coverage)
                    if len(rs1) < len(rs2):
                        temp_height = temp_temp_height
                        new_bundle = copy.deepcopy(bundle_reduced_height)
                        remaining_skus = rs1
                    elif len(rs2) < len(rs1):
                        temp_width = temp_temp_width
                        new_bundle = copy.deepcopy(bundle_reduced_width)
                        remaining_skus = rs2
                    else:
                        if height_ceiling_coverage > width_ceiling_coverage:
                            temp_height = temp_temp_height
                            new_bundle = copy.deepcopy(bundle_reduced_height)
                            remaining_skus = rs1
                        else:
                            temp_width = temp_temp_width
                            new_bundle = copy.deepcopy(bundle_reduced_width)
                            remaining_skus = rs2
                    continue

            if (bundle.height > bundle.width and
                len(bundle.skus) > 1 and
                all([sku.y == 0 for sku in bundle.skus])):
                # make sure we get unique SKUs to make sure they aren't 1 SKU stacked on itself
                unique_skus = []
                for sku in bundle.skus:
                    if sku.x not in [s.x for s in unique_skus]:
                        unique_skus.append(sku)
                    else:
                        # find SKU with max width at this x position
                        max_width_sku = max([s for s in unique_skus if s.x == sku.x], key=lambda s: s.width, default=None)
                        if max_width_sku:
                            unique_skus.remove(max_width_sku)
                            unique_skus.append(sku)
                if len(unique_skus) == 1:
                    break
                # put a filler in between SKUs
                if bundle.height < 100:
                    break
                elif bundle.height < 150:
                    filler = FILLER_44
                else:
                    filler = FILLER_62
                    filler.width, filler.height = _get_sku_dimensions(filler, True)

                middle_sku_idx = len(unique_skus) // 2
                x_to_place = unique_skus[middle_sku_idx].x
                # shift skus to the right to make space for filler
                for sku in bundle.skus:
                    if sku.x >= x_to_place:
                        sku.x += filler.width
                # place filler in the middle
                bundle.add_sku(filler, x_to_place, 0, True)
                bundle.resize_to_content()
                for sku in remaining_skus:
                    # try to place remaining SKUs in filler
                    if _place_short_sku_in_filler(bundle, sku, in_bundle=False):
                        remaining_skus.remove(sku)
            break
        # if height still larger than width, remove filler and lay flat and add board
        if bundle.height > bundle.width and bundle.skus:
            bundles.extend(_stack_skus_flat(bundle, {}))
        elif bundle.skus:
            bundles.append(bundle)

        # If no progress, skip largest SKU
        if len(remaining_skus) == before_count and remaining_skus:
            largest_sku = max(remaining_skus, key=lambda x: x.width * x.height)
            if largest_sku.id not in [sku.id for sku in REMOVED_SKUS]:
                REMOVED_SKUS.append(largest_sku)
            # remaining_skus.remove(largest_sku)
            # Create bundle with largest SKU only
            bundle_length = 3680 if (largest_sku.length < 3700) else 7340
            largest_sku.width, largest_sku.height = _get_sku_dimensions(largest_sku, False)  # un-rotate
            new_bundle = Bundle(largest_sku.width, largest_sku.height, bundle_length, packing_machine=machine)
            new_bundle.add_sku(largest_sku, 0, 0, False)  # Place SKU without rotation
            # find any stackable SKUs
            remaining_skus.remove(largest_sku)
            stackable_skus = _find_stackable_skus(largest_sku, remaining_skus, set(), -1, new_bundle.max_length, False)
            for sku in stackable_skus:
                new_bundle.add_sku(sku, 0, 0, False)
                remaining_skus.remove(sku)
            bundles.append(new_bundle)

    return bundles

def _stack_skus_flat(bundle: Bundle, sku_groups: dict = {}) -> None:
    """Lay SKUs horizontally, keeping SKU stackings and sorting by width"""
    if not sku_groups:
        # group SKUs by x position, stacks
        for sku in reversed(bundle.skus):
            if "Filler" in sku.id:
                bundle.skus.remove(sku)
                continue
            if f"{sku.x}_{sku.y}" not in sku_groups:
                sku_groups[f"{sku.x}_{sku.y}"] = []
            sku.width, sku.height = _get_sku_dimensions(sku, False) # un-rotate all SKUs
            sku.rotated = False
            sku_groups[f"{sku.x}_{sku.y}"].append(sku)

    bundles = []
    while sku_groups:
        # sort groups by length, then width of largest SKU in group
        current_y = 0
        max_width = max(sku.width for sku in bundle.skus)
        new_bundle = Bundle(max_width, bundle.height, max(sku.length for sku in bundle.skus), packing_machine=bundle.packing_machine)

        single_groups = [group for group in sku_groups if len(sku_groups[group]) == 1]
        stack_eligible_skus = []
        for group in single_groups:
            stack_eligible_skus.extend(sku_groups[group])
        empty_groups = []
        # try to find stackable SKUs for each single group
        for i, sku in enumerate(stack_eligible_skus):
            stackable_skus = _find_stackable_skus(sku, stack_eligible_skus, set(), i, new_bundle.max_length, False)
            # select the largest stackable SKU and add it to the group
            if stackable_skus:
                largest_stackable = max(stackable_skus, key=lambda s: s.length)
                for group in sku_groups.values():
                    if sku in group and id(largest_stackable) not in [id(s) for s in group]:
                        group.append(largest_stackable)
                        break
                for g_idx, group in reversed(sku_groups.items()):
                    if id(largest_stackable) in [id(s) for s in group] and len(group) == 1:
                        empty_groups.append(g_idx)
        for empty_group in list(set(empty_groups)):
            del sku_groups[empty_group]

        # sort groups by x position
        sorted_groups = sorted(sku_groups.items(), key=lambda item: (max(sku.length for sku in item[1]), max(sku.width for sku in item[1])), reverse=False)
        # try to pack each group into the bundle
        for x, group in reversed(sorted_groups):
            max_height = max(sku.height for sku in group)
            total_weight = sum(sku.weight for sku in group)
            if (current_y + max_height > max_width or new_bundle.get_total_weight() + total_weight > MAX_WEIGHT) and new_bundle.skus:
                # If adding this group exceeds bundle width or weight, stop packing and create new bundle
                new_bundle.resize_to_content()
                bundles.append(new_bundle)
                # find new max width with remaining skus in sku_groups
                max_width = max(max(sku.width for sku in sku_groups[x]) for x in sku_groups)
                current_y = 0
                new_bundle = Bundle(max_width, max_width, MAX_LENGTH, packing_machine=bundle.packing_machine)

            for sku in reversed(group):
                if current_y == 0 or _has_sufficient_support(0, current_y, sku.width, new_bundle):
                    new_bundle.add_sku(sku, 0, current_y, False)  # Place SKU without rotation
                    sku_groups[x].remove(sku)
            current_y += max_height
            if not sku_groups[x]:
                del sku_groups[x]
        if new_bundle.skus:
            new_bundle.resize_to_content()
            bundles.append(new_bundle)

    return bundles

def _pack_single_bundle(skus: List[SKU], bundle: Bundle) -> List[SKU]:
    """Pack a single bundle using vertical/horizontal pattern"""
    remaining_skus = skus.copy()
    current_y = 0

    # Set bundle max length based on available SKUs
    bundle.max_length = 3680 if max(sku.length for sku in remaining_skus if sku.length) < 3700 else 7340

    # Check for eligible bottom SKUs
    bottom_eligible_skus = [
        sku for sku in remaining_skus
        # new change: allow 289s and 145s as bottom SKUs
        if sku.can_be_bottom and sku.length > 3600 #abs(sku.length - bundle.max_length) <= 100
    ]
    if len(bottom_eligible_skus) > 0:
        # Pack bottom row first if eligible SKUs exist
        row_height = _place_bottom_row(bundle, bottom_eligible_skus, remaining_skus, True)
        current_y += row_height

    # turn all SKUs horizontal
    for sku in remaining_skus:
        sku.width, sku.height = _get_sku_dimensions(sku, False)

    short_skus = [sku for sku in remaining_skus if sku.length <= 609]
    remaining_skus = [sku for sku in remaining_skus if sku.length > 609]

    while remaining_skus and current_y < bundle.height:
        # Pack regular row
        if len(remaining_skus) <= 2 and bundle.skus:
            remaining_skus = fill_row_greedy(bundle, remaining_skus, current_y + remaining_skus[0].height-5)
            if not remaining_skus:
                break
        row_height = _pack_row(bundle, remaining_skus, current_y, bool(current_y == 0), bundle.max_length)

        if row_height == 0:
            break

        current_y += row_height
        remaining_skus = fill_row_greedy(bundle, remaining_skus, current_y)

    remaining_skus = fill_remaining_greedy(bundle, remaining_skus)

    # sort short skus
    short_skus.sort(key=lambda x: (x.width * x.height), reverse=False)
    # Add filler and check if any short SKUs can go inside filler
    # 1st pass, before short SKUs are placed
    original_width = bundle.width
    original_height = bundle.height
    bundle.resize_to_content()
    _add_filler_material(bundle)
    for sku in reversed(short_skus):
        if _place_short_sku_in_filler(bundle, sku, in_bundle=False):
            short_skus.remove(sku)
    bundle.width = original_width
    bundle.height = original_height

    # remove filler if not all short SKUs are placed
    if short_skus:
        for sku in reversed(bundle.skus):
            if "Filler" in sku.id or sku.length < 609:
                bundle.skus.remove(sku)

    if short_skus and current_y < bundle.height:
        while short_skus and current_y < bundle.height:
            # Pack short SKUs in a greedy manner
            row_height = _pack_row(bundle, short_skus, current_y, bool(current_y == 0), bundle.max_length)
            if row_height == 0:
                break

            current_y += row_height
            short_skus = fill_row_greedy(bundle, short_skus, current_y)

        # Fill remaining gaps after initial packing
        short_remaining_skus = fill_remaining_greedy(bundle, short_skus)
        remaining_skus += short_remaining_skus

    # Add filler and shrink bundle to content
    bundle.resize_to_content()
    _add_filler_material(bundle)
    # 2nd pass, move any short SKUs into filler if possible
    for sku in bundle.skus:
        _place_short_sku_in_filler(bundle, sku, in_bundle = True)

    return remaining_skus

def _place_bottom_row(bundle: Bundle, bottom_eligible_skus: List[SKU], remaining_skus: List[SKU], is_vertical_row: bool) -> int:
    """Place eligible bottom SKUs horizontally in the first row"""
    global BOTTOM_ROW_LENGTH
    for sku in bottom_eligible_skus:
        sku.width, sku.height = _get_sku_dimensions(sku, is_vertical_row)
    freq = Counter(sku.id for sku in bottom_eligible_skus)
    bottom_eligible_skus.sort(key=lambda s: (s.length, freq[s.id] * s.width, freq[s.id], s.height), reverse=True)

    current_x = 0
    row_height = 0
    row_skus = []
    remove_flag = False

    for i, sku in enumerate(bottom_eligible_skus[:]):
        if remove_flag:
            bottom_eligible_skus.remove(sku)
            remaining_skus.remove(sku)
            remove_flag = False
            continue

        if (current_x + sku.width <= bundle.width and
            (row_height == 0 or sku.height <= row_height) and
            _can_fit_in_bundle(sku, current_x, 0, is_vertical_row, bundle) and
            _sku_within_height_range(sku, row_skus) and
            sum(s[1].weight for s in row_skus) + sku.weight <= MAX_WEIGHT):

            row_skus.append((i, sku, current_x, 0, is_vertical_row))
            if sku.length == 3650:
                for test_sku in bottom_eligible_skus:
                    if test_sku.id == sku.id and (sku.length + test_sku.length <= bundle.max_length) and id(test_sku) != id(sku):
                        row_skus.append((i, test_sku, current_x, 0, is_vertical_row))
                        remove_flag = True
                        break
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
    # sort remaining SKUs by how many appear in the list + bundle
    freq = Counter(sku.id for sku in (remaining_skus + bundle.skus))
    remaining_skus.sort(key=lambda s: (freq[s.id] * s.width, freq[s.id], s.width), reverse=True)

    # Get accurate bundle weights
    total_weight = bundle.get_total_weight()
    running_total = 0
    for i, sku in enumerate(remaining_skus):
        bundle_weight = total_weight + running_total + sku.weight
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
            _can_fit_in_bundle(sku, current_x, current_y, is_vertical_row, bundle) and
            # Check if adding this SKU would exceed bundle weight limit
            bundle_weight <= MAX_WEIGHT):

            # Check support for non-bottom rows
            if current_y != 0 and not _has_sufficient_support(current_x, current_y, width, bundle):
                continue

            # Find stackable SKUs
            stackable_skus = _find_stackable_skus(sku, remaining_skus, considered_skus, i, max_length, is_vertical_row)

            # Mark all SKUs in this stack as considered
            considered_skus.add(id(sku))
            for stackable_sku in reversed(stackable_skus):
                if stackable_sku.weight + bundle_weight <= MAX_WEIGHT:
                    considered_skus.add(id(stackable_sku))
                else:
                    stackable_skus.remove(stackable_sku)

            row_skus.append((i, sku, current_x, current_y, width, height, stackable_skus))
            running_total += sku.weight + (max(stack_sku.weight for stack_sku in stackable_skus) if stackable_skus else 0)
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
                if bundle.max_length == 7340:
                    # Add a second filler for 7340 length bundles
                    bundle.add_sku(filler_copy, x, y, rotated)
                placed_any = True
                break

def _place_short_sku_in_filler(bundle: Bundle, sku, in_bundle: bool) -> None:
    """Place short SKUs inside filler material if possible"""
    if not sku or sku.length > 609:
        return False

    # Try to place the short SKU in existing filler
    for filler in bundle.skus:
        if "Filler" not in filler.id:
            continue

        # Check if SKU can fit inside this filler
        for rotated in [False, True]:
            width, height = _get_sku_dimensions(sku, rotated)

            if (width <= filler.width + 1 and height <= filler.height + 1 and sku.length <= filler.length):
                # _can_place_sku_at_position(sku, filler.x, filler.y, sku.width, sku.height, bundle)):
                sku.width, sku.height = width, height
                if in_bundle:
                    # just move sku to filler position
                    sku.x, sku.y = filler.x, filler.y
                    sku.rotated = rotated
                else:
                    bundle.add_sku(sku, filler.x, filler.y, False)
                return True

    return False

# Helper functions

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
                sku.width, sku.height = _get_sku_dimensions(sku, rotated)

                # Skip if too big for bundle
                if sku.width > bundle.width or sku.height > bundle.height:
                    continue

                for (x, y) in candidate_points:
                    if x + sku.width > bundle.width or y + sku.height > bundle.height or y + sku.height > BOTTOM_ROW_LENGTH:
                        continue

                    if _can_place_sku_at_position(sku, x, y, sku.width, sku.height, bundle):
                        # Check support if not on bottom
                        if ((y > 0 and not _has_sufficient_support(x, y, sku.width, bundle)) or
                            (y == 0 and (abs(sku.length - bundle.max_length) > 100 or not sku.can_be_bottom)) or
                            (y == 0 and not rotated) or
                            (rotated and (y + sku.height > 10 + max([sku.y + sku.height for sku in bundle.skus])))):
                            continue

                        # Find stackable SKUs
                        stackable_skus = _find_stackable_skus(sku, remaining_skus, considered_skus, i, bundle.max_length, rotated)

                        # Mark all SKUs in this stack as considered
                        considered_skus.add(id(sku))
                        for stackable_sku in stackable_skus:
                            considered_skus.add(id(stackable_sku))

                        # Place the main SKU
                        bundle.add_sku(sku, x, y, rotated)

                        # Place stackable SKUs at the same position
                        for stack_sku in stackable_skus:
                            bundle.add_sku(stack_sku, x, y, rotated)

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
    y_buffer = 70
    x_minimum_for_buffer = 0.7
    placed = True
    while placed and remaining_skus:
        placed = False
        # candidate point generation
        candidate_points = set()
        for sku in bundle.skus:
            candidate_points.add((sku.x + sku.width, sku.y))
            candidate_points.add((sku.x, sku.y + sku.height))
        # grid
        for gx in range(0, bundle.width, 50):
            for gy in range(0, round(y_limit + y_buffer), 50):
                if (gy <= y_limit or
                    (gy > y_limit and gx >= bundle.width * x_minimum_for_buffer)):
                    candidate_points.add((gx, gy))

        # Track which SKUs we've already considered for stacking
        considered_skus = set()

        for i, sku in enumerate(sorted(remaining_skus, key=lambda s: s.width*s.height, reverse=True)):
            if id(sku) in considered_skus:
                continue

            for rot in (False, True):
                w, h = _get_sku_dimensions(sku, rot)
                if w > bundle.width or h > bundle.height or h > y_limit + y_buffer:
                    continue
                for x,y in sorted(candidate_points, key=lambda p: (p[1], p[0])):
                    if (x+w > bundle.width or
                        y+h > y_limit + y_buffer or
                        (y + h > y_limit and x < bundle.width * x_minimum_for_buffer and y != 0) or
                        (y == 0 and abs(sku.length - bundle.max_length) > 100) or
                        (y == 0 and not rot)
                        # (rot and (y + h > bundle.height))
                    ):
                        continue
                    if _can_place_sku_at_position(sku, x, y, w, h, bundle) and \
                        ((y == 0 and sku.can_be_bottom) or _has_sufficient_support(x, y, w, bundle)):

                        x_shift = x
                        # Move x position left as much as possible
                        if y == 0 and x > 0:
                            while (x_shift > 0 and
                                   _can_place_sku_at_position(sku, x_shift - 5, y, w, h, bundle)):
                                x_shift -= 5
                            x = x_shift

                        # Find stackable SKUs
                        original_index = remaining_skus.index(sku)
                        stackable_skus = _find_stackable_skus(sku, remaining_skus, considered_skus, original_index, bundle.max_length, rot)

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

def _process_override_bundles(skus: List[SKU], bundle_width: int, bundle_height: int, mach1_skus: List[SKU]) -> List[Bundle]:
    """Process SKUs with bundle override"""
    if not skus:
        return []
    override_groups = {}
    for sku in skus:
        override_id = sku.data.get('Bdl_Override')
        if override_id not in override_groups:
            override_groups[override_id] = []
        override_groups[override_id].append(sku)

    bundles = []
    for override_skus in override_groups.values():
        # return an error if MACH1 and MACH5 skus are mixed
        color_groups = _group_skus_by_color(override_skus)
        color_groups = [col[-3:] for col in color_groups.keys()]
        if 0 < len(set(color_groups).intersection(set(mach1_skus))) < len(color_groups):
            return -1
        machine = "MACH1" if any(sku.id[-3:] in mach1_skus for sku in override_skus) else "MACH5"
        override_skus.sort(key=lambda x: max(x.height, x.width), reverse=True)

        current_bundles = _pack_skus_with_pattern(override_skus, bundle_width, bundle_height, machine=machine)
        # attempt to merge in case of any missed space
        merged_bundles = _try_merge_bundles(current_bundles, bundle_width, bundle_height, machine=machine)
        for bundle in merged_bundles:
            if bundle.skus:
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
    if x + width > bundle.width or y + height > bundle.height or sku.weight + bundle.get_total_weight() > MAX_WEIGHT:
        return False
    if y == 0 and (not sku.can_be_bottom):# or (bundle.max_length == 7340 and sku.length < 3700)):
        return False

    for placed_sku in bundle.skus:
        if (x < placed_sku.x + placed_sku.width and
            x + width > placed_sku.x and
            y < placed_sku.y + placed_sku.height and
            y + height > placed_sku.y):
            return False
    return True

def _has_sufficient_ceiling_coverage(bundle: Bundle, get_value: bool = False) -> bool:
    """Check if the bundle has sufficient coverage along the top of the bundle"""
    copy_bundle = copy.deepcopy(bundle)
    copy_bundle.resize_to_content()
    _add_filler_material(copy_bundle)
    buffer = MAX_DIST_FROM_CEILING
    required_coverage = MIN_CEILING_COVERAGE # % coverage required

    if not copy_bundle.skus:
        return False
    support_segments = []
    for sku in copy_bundle.skus:
        if sku.y + sku.height >= copy_bundle.height - buffer:
            overlap_start = max(0, sku.x)
            overlap_end = min(copy_bundle.width, sku.x + sku.width)
            if overlap_end > overlap_start:
                support_segments = list(set(support_segments).union(list(range(round(overlap_start), round(overlap_end)))))

    total_coverage = len(support_segments)
    if get_value:
        return total_coverage / (copy_bundle.width)
    return total_coverage >= copy_bundle.width * required_coverage

def _has_sufficient_support(x: int, y: int, width: int, bundle: Bundle, get_value: bool = False) -> bool:
    """Check if position has sufficient support from SKUs below"""
    threshold = BASE_COVERAGE_THRESHOLD
    buffer = SKU_COVERAGE_HEIGHT_BUFFER
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

def _find_stackable_skus(target_sku: SKU, remaining_skus: List[SKU], unavailable_skus: set, target_index: int, max_length: int, rotated: bool) -> List[SKU]:
    """Find SKUs that can be stacked with target SKU based on combined length"""
    stackable = []
    current_total_length = target_sku.length
    
    # Create a list of candidates with their indices for sorting
    candidates = []
    for i, sku in enumerate(remaining_skus):
        if id(sku) in unavailable_skus:
            continue
        sku.width, sku.height = _get_sku_dimensions(sku, rotated)
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
    return (abs(sku1.width - sku2.width) <= STACKING_MAX_DIFF and
            abs(sku1.height - sku2.height) <= STACKING_MAX_DIFF)

def _sku_within_height_range(sku: SKU, row_skus: List) -> bool:
    """Check if SKU is within height tolerance of previous SKU"""
    height_tol = SKU_MAX_HEIGHT_DIFF # mm tolerance for height matching
    if not row_skus:
        return True
    last_sku_height = row_skus[-1][1].height
    return (last_sku_height - height_tol <= sku.height <= last_sku_height + height_tol)