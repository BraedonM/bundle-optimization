from dataclasses import dataclass
from typing import List, Tuple, Optional

@dataclass
class SKU:
    id: str
    width: float
    height: float
    length: float
    weight: float
    desc: str


@dataclass
class PlacedSKU(SKU):
    x: int
    y: int
    rotated: bool = False


@dataclass
class Bundle:
    width: float
    height: float
    skus: List[PlacedSKU]
    free_rects: List[Tuple[int, int, float, float]]  # x, y, w, h

    def __init__(self, width: float, height: float):
        self.width = width
        self.height = height
        self.skus = []
        self.free_rects = [(0, 0, width, height)]

    def try_place_sku(self, sku: SKU) -> Optional[PlacedSKU]:
        best_score = None
        best_rect = None
        best_rotated = False

        options = [
            (sku.width, sku.height, False),
            (sku.height, sku.width, True)
        ]

        for w, h, rotated in options:
            for rect in self.free_rects:
                rx, ry, rw, rh = rect
                if w <= rw and h <= rh:
                    leftover_h = rh - h
                    leftover_w = rw - w
                    short_side_fit = min(leftover_w, leftover_h)
                    long_side_fit = max(leftover_w, leftover_h)

                    score = (short_side_fit, long_side_fit)

                    if best_score is None or score < best_score:
                        best_score = score
                        best_rect = (rx, ry, w, h)
                        best_rotated = rotated

        if best_rect is not None:
            x, y, w, h = best_rect
            placed = PlacedSKU(sku.id, w, h, sku.length, sku.weight, sku.desc, x, y, best_rotated)
            self.skus.append(placed)
            self.split_free_rects(x, y, w, h)
            self.prune_free_list()
            return placed

        return None

    def add_filler_materials(self):
        """
        Try to add filler materials to remaining free space in the bundle.
        Filler materials:
        - Pack_44Filler: 100x100mm, 1.810kg
        - Pack_62Filler: 150x50mm, 2.268kg
        """
        # Define filler materials
        fillers = [
            SKU("Pack_44Filler", 100, 100, 100, 1.810, "Filler Material 100x100mm"),
            SKU("Pack_62Filler", 150, 50, 50, 2.268, "Filler Material 150x50mm")
        ]

        # Keep trying to place fillers until no more can be placed
        added_any = True
        while added_any:
            added_any = False

            # Try each filler type
            for filler in fillers:
                # Keep trying to place this filler type until it can't fit anymore
                while True:
                    placed_filler = self.try_place_sku(filler)
                    if placed_filler is None:
                        break
                    added_any = True

    def split_free_rects(self, x: int, y: int, w: float, h: float):
        new_rects = []
        for rx, ry, rw, rh in self.free_rects:
            if x + w <= rx or x >= rx + rw or y + h <= ry or y >= ry + rh:
                # No overlap
                new_rects.append((rx, ry, rw, rh))
                continue

            # Split logic (guillotine-style)
            if x > rx:
                new_rects.append((rx, ry, x - rx, rh))
            if x + w < rx + rw:
                new_rects.append((x + w, ry, (rx + rw) - (x + w), rh))
            if y > ry:
                new_rects.append((rx, ry, rw, y - ry))
            if y + h < ry + rh:
                new_rects.append((rx, y + h, rw, (ry + rh) - (y + h)))

        self.free_rects = new_rects

    def prune_free_list(self):
        # Remove any rects that are fully contained in others
        pruned = []
        for i, r in enumerate(self.free_rects):
            contained = False
            for j, r2 in enumerate(self.free_rects):
                if i != j and self.is_contained_in(r, r2):
                    contained = True
                    break
            if not contained:
                pruned.append(r)
        self.free_rects = pruned

    @staticmethod
    def is_contained_in(r1, r2):
        x1, y1, w1, h1 = r1
        x2, y2, w2, h2 = r2
        return x1 >= x2 and y1 >= y2 and x1 + w1 <= x2 + w2 and y1 + h1 <= y2 + h2
