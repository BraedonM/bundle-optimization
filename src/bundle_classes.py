from dataclasses import dataclass
from typing import List

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
    stacked_quantity: int = 1  # Number of SKUs stacked in this position

@dataclass
class Bundle:
    width: float
    height: float
    max_length: float = 3680  # Maximum length for the bundle
    skus: List[PlacedSKU] = None

    def __init__(self, width: float, height: float, max_length: float = 3680):
        self.width = width
        self.height = height
        self.max_length = max_length
        self.skus = []

    def add_sku(self, sku: SKU, x: int, y: int, rotated: bool, stacked_quantity: int = 1) -> PlacedSKU:
        """
        Places an SKU at a specific location without validation.
        The caller is responsible for ensuring it fits.
        """
        placed = PlacedSKU(
            id=sku.id,
            width=sku.width,
            height=sku.height,
            length=sku.length,
            weight=sku.weight,
            desc=sku.desc,
            x=x,
            y=y,
            rotated=rotated,
            stacked_quantity=stacked_quantity
        )
        self.skus.append(placed)
        return placed

    def get_actual_dimensions(self):
        """
        Calculate the actual dimensions of the bundle based on placed SKUs
        """
        if not self.skus:
            return 0, 0, 0
        
        max_x = max(sku.x + sku.width for sku in self.skus)
        max_y = max(sku.y + sku.height for sku in self.skus)
        max_length = max(sku.length for sku in self.skus)
        
        return max_x, max_y, max_length

    def get_total_weight(self):
        """
        Calculate total weight including stacked quantities and packaging materials
        """
        return self.add_package_weight() + sum(sku.weight * sku.stacked_quantity for sku in self.skus)

    def resize_to_content(self):
        """
        Resize bundle to fit the actual content
        """
        if self.skus:
            actual_width, actual_height, actual_length = self.get_actual_dimensions()
            self.width = actual_width
            self.height = actual_height
            self.max_length = actual_length

    def add_package_weight(self):
        """
        Add the weight from packaging material to the bundle
        """
        width = self.width
        height = self.height
        weight = 0

        # Add weights
        if self.max_length <= 3700: # 3680mm with some error
            self.max_length = 3680
            weight += 11 # 11kg of packaging weight

            # Pack pad weight
            if width <= 216: # 8.5 inches
                weight += 0.898
            elif width <= 254: # 10 inches
                weight += 1.995
            elif width <= 331: # 13 inches
                weight += 1.995
            else: # 19 inches
                weight += 2.721

            if height <= 216: # 8.5 inches
                weight += 0.898
            elif height <= 254: # 10 inches
                weight += 1.995
            elif height <= 331: # 13 inches
                weight += 1.995
            else: # 19 inches
                weight += 2.721

        else: # 7340mm
            self.max_length = 7340
            weight += 22 # 22kg of packaging weight

            # Pack pad weight
            if width <= 216: # 8.5 inches
                weight += 1.796
            elif width <= 254: # 10 inches
                weight += 3.991
            elif width <= 331: # 13 inches
                weight += 3.991
            else: # 19 inches
                weight += 5.442

            if height <= 216: # 8.5 inches
                weight += 1.796
            elif height <= 254: # 10 inches
                weight += 3.991
            elif height <= 331: # 13 inches
                weight += 3.991
            else: # 19 inches
                weight += 5.442

        return weight


