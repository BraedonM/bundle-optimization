import matplotlib.pyplot as plt
import matplotlib.patches as patches
import random
from typing import List

from bundle_classes import Bundle


def visualize_bundles(bundles: List[Bundle], savePath: str = None) -> None:
    num_bundles = len(bundles)
    fig, axs = plt.subplots(1, num_bundles, figsize=(6 * num_bundles, 6))

    sku_colors = {}

    if num_bundles == 1:
        axs = [axs]

    for idx, (bundle, ax) in enumerate(zip(bundles, axs)):
        ax.set_title(f"Bundle {idx + 1}")
        ax.set_xlim(0, bundle.width)
        ax.set_ylim(0, bundle.height)
        ax.set_aspect('equal')
        ax.set_xticks([])
        ax.set_yticks([])
        ax.grid(False)

        for sku in bundle.skus:
            if sku.id not in sku_colors:
                sku_colors[sku.id] = [random.random() * 0.7 + 0.3 for _ in range(3)]
            color = sku_colors[sku.id]
            rect = patches.Rectangle(
                (sku.x, sku.y),
                sku.width,
                sku.height,
                linewidth=1,
                edgecolor='black',
                facecolor=color,
            )
            ax.add_patch(rect)

            # Only label if the area is large enough
            if sku.width > 15 and sku.height > 10:
                label_x = sku.x + sku.width / 2
                label_y = sku.y + sku.height / 2
                ax.text(label_x, label_y, sku.id, ha='center', va='center', fontsize=7, weight='bold')

    plt.tight_layout()
    if savePath:
        plt.savefig(savePath, dpi=300)
        plt.close(fig)
    else:
        print("Path not found, displaying the plot instead.")
        plt.show()