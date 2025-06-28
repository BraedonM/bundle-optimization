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

@dataclass
class Bundle:
    width: float
    height: float
    skus: List[PlacedSKU]

    def __init__(self, width: float, height: float):
        self.width = width
        self.height = height
        self.skus = []

    def add_sku(self, sku: SKU, x: int, y: int, rotated: bool) -> PlacedSKU:
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
            rotated=rotated
        )
        self.skus.append(placed)
        return placed
