import matplotlib.pyplot as plt
import matplotlib.patches as patches
import random
from typing import List
import numpy as np
import copy

from bundle_classes import Bundle


def visualize_bundles(original_bundles: List[Bundle], savePath: str = None, unit: str = 'metric',
                      packaging_height: float = 0, packaging_width: float = 0, lumber_height: float = 0) -> None:
    bundles = copy.deepcopy(original_bundles)
    num_bundles = len(bundles)
    try:
        fig, axs = plt.subplots(1, num_bundles, figsize=(6 * num_bundles, 6))
    except ValueError:
        return

    sku_colors = {}

    if num_bundles == 1:
        axs = [axs]

    for idx, (bundle, ax) in enumerate(zip(bundles, axs)):
        # Sort skus by size
        bundle.skus.sort(key=lambda sku: (sku.width * sku.height), reverse=True)
        # Use actual bundle dimensions for visualization

        weight_kg = bundle.get_total_weight()
        # remove packaging skus from bundle so they don't show up in the visualization
        bundle.skus = [sku for sku in bundle.skus if (not sku.id.startswith("Pack_") or "Filler" in sku.id)]
        actual_width, actual_height, max_length = bundle.get_actual_dimensions()
        lumber = lumber_height if all([sku.rotated is False for sku in bundle.skus]) else 0

        display_width = actual_width + packaging_width
        display_height = actual_height + packaging_height + lumber

        if unit == 'imperial':
            actual_width /= 25.4
            actual_height /= 25.4
            display_width /= 25.4
            display_height /= 25.4
            max_length /= 25.4
            weight_kg *= 2.20462
            weight_unit = 'lbs'
            length_unit = 'in'
            ticks = 2
            length_divisor = 25.4
        else:
            weight_unit = 'kg'
            length_unit = 'mm'
            ticks = 50
            length_divisor = 1

        ax.set_title(f"Bundle {idx + 1}\n({display_width:.0f}x{display_height:.0f}x{max_length:.0f}{length_unit}, {weight_kg:.2f}{weight_unit})")
        ax.set_xlim(0, actual_width)
        ax.set_ylim(0, actual_height)
        ax.set_aspect('equal')
        # set ticks every 25mm
        if bundle.height > 400 or bundle.width > 400:
            ax.set_xticks(np.arange(0, actual_width, ticks))
            ax.set_yticks(np.arange(0, actual_height, ticks))
        else:
            ax.set_xticks(np.arange(0, actual_width, ticks/2))
            ax.set_yticks(np.arange(0, actual_height, ticks/2))
        ax.grid(False)

        sku_locations = {}
        for sku in bundle.skus:
            # Convert SKU to accurate unit
            sku.x /= length_divisor
            sku.y /= length_divisor
            sku.width /= length_divisor
            sku.height /= length_divisor
            loc_key = (sku.x, sku.y)
            if loc_key not in sku_locations:
                sku_locations[loc_key] = []
            sku_locations[loc_key].append(sku)

        for sku in bundle.skus:
            # Assign consistent colors, with different shades for filler materials
            if sku.id not in sku_colors:
                if "Filler" in sku.id:
                    # Use gray tones for filler materials
                    sku_colors[sku.id] = [0.7, 0.7, 0.7]  # Light gray
                elif "Partial" in sku.id:
                    # Use same color as original SKU
                    sku_colors[sku.id] = sku_colors[sku.id.replace("_Partial", "")] if sku.id.replace("_Partial", "") in sku_colors else [random.random() * 0.7 + 0.3 for _ in range(3)]
                else:
                    # Use bright colors for regular SKUs
                    sku_colors[sku.id] = [random.random() * 0.7 + 0.3 for _ in range(3)]

            color = sku_colors[sku.id]

            # Create rectangle with different border style for filler
            if "Filler" in sku.id:
                rect = patches.Rectangle(
                    (sku.x, sku.y),
                    sku.width,
                    sku.height,
                    linewidth=2,
                    edgecolor='red',
                    facecolor=color,
                    linestyle='--',  # Dashed line for filler
                    alpha=0.7
                )
            else:
                rect = patches.Rectangle(
                    (sku.x, sku.y),
                    sku.width,
                    sku.height,
                    linewidth=1,
                    edgecolor='black',
                    facecolor=color,
                )
            ax.add_patch(rect)

        for loc_key, skus in sku_locations.items():
            larger_sku = max(skus, key=lambda s: s.width * s.height)
            # Label with SKU ID and quantity if stacked
            label_x = larger_sku.x + larger_sku.width / 2
            label_y = larger_sku.y + larger_sku.height / 2
            skus_same = {}
            for sku in skus:
                if sku.height == 0 or sku.width == 0:
                    continue
                if sku.id not in skus_same:
                    skus_same[sku.id] = 0
                skus_same[sku.id] += 1

            label_text = ""
            for sku_id, quantity in skus_same.items():
                if "Partial" in sku_id:
                    sku_id = sku_id.replace("_Partial", "\n(Partial)")
                if label_text:
                    label_text += "\n"
                if quantity > 1:
                    label_text += f"{sku_id}\n(x{quantity})"
                else:
                    label_text += sku_id

            ax.text(label_x, label_y, label_text, ha='center', va='center',
                   fontsize=4, weight='bold',
                   bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.8))

    plt.tight_layout()
    if savePath:
        plt.savefig(savePath, dpi=300, bbox_inches='tight')
        plt.close(fig)
    else:
        print("Path not found, displaying the plot instead.")
        plt.show()
