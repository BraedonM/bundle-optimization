# bundle_classes.py
from dataclasses import dataclass
from typing import List

@dataclass
class SKU:
    id: str
    bundleqty: int = 1  # Number of SKUs in a bundle
    width: float = 0
    height: float = 0
    length: float = 0
    weight: float = 0
    desc: str = ''
    can_be_bottom: bool = False  # Can this SKU be placed at the bottom of a bundle
    data : dict = None  # Additional data for SKU that will not be changed

@dataclass
class PlacedSKU(SKU):
    x: int = 0
    y: int = 0
    rotated: bool = False

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

    def add_sku(self, sku: SKU, x: int, y: int, rotated: bool) -> PlacedSKU:
        """
        Places an SKU at a specific location without validation.
        The caller is responsible for ensuring it fits.
        """
        placed = PlacedSKU(
            # SKU properties
            id=sku.id,
            bundleqty=sku.bundleqty,
            width=sku.width,
            height=sku.height,
            length=sku.length,
            weight=sku.weight,
            desc=sku.desc,
            can_be_bottom=sku.can_be_bottom,
            data=sku.data,

            # PlacedSKU properties
            x=x,
            y=y,
            rotated=rotated,
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
        # max_length = max(sku.length for sku in self.skus)
        max_length = 3680 if (max(sku.length for sku in self.skus if sku.length) < 3700) else 7340

        return max_x, max_y, max_length

    def get_total_weight(self):
        """
        Calculate total weight including packaging materials
        """
        return sum(sku.weight for sku in self.skus)

    def resize_to_content(self):
        """
        Resize bundle to fit the actual content
        """
        if self.skus:
            actual_width, actual_height, actual_length = self.get_actual_dimensions()
            self.width = actual_width
            self.height = actual_height
            self.max_length = actual_length

    def add_packaging(self):
        """
        Add the SKUs from packaging material to the bundle
        """
        width, height, actual_length = self.get_actual_dimensions()

        # Check if all SKUs are horizontal (need boards)
        add_lumber = False
        if all([sku.rotated is False for sku in self.skus]):
            add_lumber = True

        # Add weights
        if actual_length == 3680:
            self.add_sku(PACK_ANGLE_3680, 0, 0, False)
            self.add_sku(PACK_1_4_19_DUN_3680, 0, 0, False)
            self.add_sku(PACK_2_3_19_DUN_3680, 0, 0, False)
            self.add_sku(PACK_SUB_BUNDL_WRP_3680, 0, 0, False)
            self.add_sku(PACK_MST_BUNDL_WRP_3680, 0, 0, False)

            self.max_length = 3680

            # Pack pad weight
            if width <= 216: # 8.5 inches
                self.add_sku(PACK_PAD_8_3680, 0, 0, False)
            elif width <= 254: # 10 inches
                self.add_sku(PACK_PAD_10_3680, 0, 0, False)
            elif width <= 331: # 13 inches
                self.add_sku(PACK_PAD_13_3680, 0, 0, False)
            else: # 19 inches
                self.add_sku(PACK_PAD_19_3680, 0, 0, False)

            if height <= 216: # 8.5 inches
                self.add_sku(PACK_PAD_8_3680, 0, 0, False)
            elif height <= 254: # 10 inches
                self.add_sku(PACK_PAD_10_3680, 0, 0, False)
            elif height <= 331: # 13 inches
                self.add_sku(PACK_PAD_13_3680, 0, 0, False)
            else: # 19 inches
                self.add_sku(PACK_PAD_19_3680, 0, 0, False)

            if add_lumber:
                self.add_sku(PACK_LUMBER_3680, 0, 0, False)
                if height > 100:
                    self.add_sku(PACK_LUMBER_3680, 0, 0, False)

        else: # 7340mm
            self.add_sku(PACK_ANGLE_7340, 0, 0, False)
            self.add_sku(PACK_1_4_19_DUN_7340, 0, 0, False)
            self.add_sku(PACK_2_3_19_DUN_7340, 0, 0, False)
            self.add_sku(PACK_SUB_BUNDL_WRP_7340, 0, 0, False)
            self.add_sku(PACK_MST_BUNDL_WRP_7340, 0, 0, False)

            self.max_length = 7340

            # Pack pad weight
            if width <= 216: # 8.5 inches
                self.add_sku(PACK_PAD_8_7340, 0, 0, False)
            elif width <= 254: # 10 inches
                self.add_sku(PACK_PAD_10_7340, 0, 0, False)
            elif width <= 331: # 13 inches
                self.add_sku(PACK_PAD_13_7340, 0, 0, False)
            else: # 19 inches
                self.add_sku(PACK_PAD_19_7340, 0, 0, False)

            if height <= 216: # 8.5 inches
                self.add_sku(PACK_PAD_8_7340, 0, 0, False)
            elif height <= 254: # 10 inches
                self.add_sku(PACK_PAD_10_7340, 0, 0, False)
            elif height <= 331: # 13 inches
                self.add_sku(PACK_PAD_13_7340, 0, 0, False)
            else: # 19 inches
                self.add_sku(PACK_PAD_19_7340, 0, 0, False)

            if add_lumber:
                self.add_sku(PACK_LUMBER_7340, 0, 0, False)
                if height > 100:
                    self.add_sku(PACK_LUMBER_7340, 0, 0, False)
        return

## PREDEFINED SKUs for Filler and Packaging
# Filler materials
FILLER_44 = SKU(
    id="Pack_44Filler",
    bundleqty=1,
    width=100,
    height=100,
    length=3660,
    weight=1.810,
    desc="Pack 44 Filler Material"
)

FILLER_62 = SKU(
    id="Pack_62Filler",
    bundleqty=1,
    width=150,
    height=50,
    length=3660,
    weight=2.268,
    desc="Pack 62 Filler Material"
)

# Packaging SKUs
PACK_ANGLE_3680 = SKU(
    id="Pack_Angle",
    bundleqty=4,
    length=3660,
    weight=5.442,
    desc="PRINTED ANGLEBOARD 3680mm"
)

PACK_ANGLE_7340 = SKU(
    id="Pack_Angle",
    bundleqty=8,
    length=7320,
    weight=10.884,
    desc="PRINTED ANGLEBOARD 7340mm"
)

PACK_1_4_19_DUN_3680 = SKU(
    id="Pack_1x4x19_Dun",
    bundleqty=2,
    length=482.6,
    weight=0.998,
    desc="1\" X 4\" X 19\" DUNNAGE 3680mm"
)

PACK_1_4_19_DUN_7340 = SKU(
    id="Pack_1x4x19_Dun",
    bundleqty=4,
    length=965.2,
    weight=1.995,
    desc="1\" X 4\" X 19\" DUNNAGE 7340mm"
)

PACK_2_3_19_DUN_3680 = SKU(
    id="Pack_2x3x19_Dun",
    bundleqty=2,
    length=482.6,
    weight=1.796,
    desc="2\" X 3\" X 19\" DUNNAGE 3680mm"
)

PACK_2_3_19_DUN_7340 = SKU(
    id="Pack_2x3x19_Dun",
    bundleqty=4,
    length=965.2,
    weight=3.592,
    desc="2\" X 3\" X 19\" DUNNAGE 7340mm"
)

PACK_LUMBER_3680 = SKU(
    id="Pack_2x3x19_Dun_Lumber_3680",
    bundleqty=1,
    length=3660,
    weight=2.721,
    desc="COMMON LUMBER - 1\" X 4\" X 12\" 3680mm"
)

PACK_LUMBER_7340 = SKU(
    id="Pack_2x3x19_Dun_Lumber_7340",
    bundleqty=2,
    length=7320,
    weight=5.442,
    desc="COMMON LUMBER - 1\" X 4\" X 12\" 7340mm"
)

PACK_PAD_8_3680 = SKU(
    id="Pack_Pad_8_3600",
    bundleqty=2,
    length=3660,
    weight=0.898,
    desc="PAD - 8-1/2\" X 144\" DW ECT #3 WHITE 3680mm"
)

PACK_PAD_8_7340 = SKU(
    id="Pack_Pad_8_7340",
    bundleqty=4,
    length=7320,
    weight=1.796,
    desc="PAD - 8-1/2\" X 144\" DW ECT #3 WHITE 7340mm"
)

PACK_PAD_10_3680 = SKU(
    id="Pack_Pad_10_3680",
    bundleqty=2,
    length=3660,
    weight=1.995,
    desc="PAD - 10\" X 144\" DW ECT #3 WHITE 3680mm"
)

PACK_PAD_10_7340 = SKU(
    id="Pack_Pad_10_7340",
    bundleqty=4,
    length=7320,
    weight=3.991,
    desc="PAD - 10\" X 144\" DW ECT #3 WHITE 7340mm"
)

PACK_PAD_13_3680 = SKU(
    id="Pack_Pad_13_3680",
    bundleqty=2,
    length=3660,
    weight=1.995,
    desc="PAD - 13\" X 144\" DW ECT #3 WHITE 3680mm"
)

PACK_PAD_13_7340 = SKU(
    id="Pack_Pad_13_7340",
    bundleqty=4,
    length=7320,
    weight=3.991,
    desc="PAD - 13\" X 144\" DW ECT #3 WHITE 7340mm"
)

PACK_PAD_19_3680 = SKU(
    id="Pack_Pad_19_3680",
    bundleqty=2,
    length=3660,
    weight=2.721,
    desc="PAD - 19\" X 144\" DW ECT #3 WHITE 3680mm"
)

PACK_PAD_19_7340 = SKU(
    id="Pack_Pad_19_7340",
    bundleqty=4,
    length=7320,
    weight=5.442,
    desc="PAD - 19\" X 144\" DW ECT #3 WHITE 7340mm"
)

PACK_SUB_BUNDL_WRP_3680 = SKU(
    id="Pack_Sub_Bundl_Wrp_3680",
    bundleqty=2,
    length=3660,
    weight=2.268,
    desc="Sub-Bundle Wrap - Crepe Paper/Stretch Film 3680mm"
)

PACK_SUB_BUNDL_WRP_7340 = SKU(
    id="Pack_Sub_Bundl_Wrp_7340",
    bundleqty=4,
    length=7320,
    weight=4.535,
    desc="Sub-Bundle Wrap - Crepe Paper/Stretch Film 7340mm"
)

PACK_MST_BUNDL_WRP_3680 = SKU(
    id="Pack_Mst_Bundl_Wrp_3680",
    bundleqty=2,
    length=3660,
    weight=0.500,
    desc="Master Bundle - Stretch Wrap 3680mm"
)

PACK_MST_BUNDL_WRP_7340 = SKU(
    id="Pack_Mst_Bundl_Wrp_7340",
    bundleqty=4,
    length=7320,
    weight=1.000,
    desc="Master Bundle - Stretch Wrap 7340mm"
)