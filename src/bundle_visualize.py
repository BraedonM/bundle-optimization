import matplotlib.pyplot as plt
import matplotlib.patches as patches
import random
from typing import List

from bundle_classes import Bundle


def visualize_bundles(bundles: List[Bundle], savePath: str = None) -> None:
    num_bundles = len(bundles)
    try:
        fig, axs = plt.subplots(1, num_bundles, figsize=(6 * num_bundles, 6))
    except ValueError:
        return

    sku_colors = {}

    if num_bundles == 1:
        axs = [axs]

    for idx, (bundle, ax) in enumerate(zip(bundles, axs)):
        # Use actual bundle dimensions for visualization
        actual_width, actual_height, max_length = bundle.get_actual_dimensions()

        ax.set_title(f"Bundle {idx + 1}\n({actual_width:.0f}x{actual_height:.0f}x{max_length:.0f}mm, {bundle.get_total_weight():.2f}kg)")
        ax.set_xlim(0, actual_width)
        ax.set_ylim(0, actual_height)
        ax.set_aspect('equal')
        ax.set_xticks([])
        ax.set_yticks([])
        ax.grid(False)
        
        # Group SKUs by their x and y coordinates
        sku_locations = {}
        for sku in bundle.skus:
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
