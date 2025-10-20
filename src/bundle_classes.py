# bundle_classes.py
from dataclasses import dataclass
from typing import List

# Packaging and filler materials
PACK_ANGLE_3680 = None
PACK_ANGLE_7340 = None
PACK_1_4_19_DUN_3680 = None
PACK_1_4_19_DUN_7340 = None
PACK_2_3_19_DUN_3680 = None
PACK_2_3_19_DUN_7340 = None
PACK_LUMBER_3680 = None
PACK_LUMBER_7340 = None
PACK_PAD_8_3680 = None
PACK_PAD_8_7340 = None
PACK_PAD_10_3680 = None
PACK_PAD_10_7340 = None
PACK_PAD_13_3680 = None
PACK_PAD_13_7340 = None
PACK_PAD_19_3680 = None
PACK_PAD_19_7340 = None
PACK_SUB_BNDL_WRP_3680 = None
PACK_SUB_BNDL_WRP_7340 = None
PACK_MST_BNDL_WRP_3680 = None
PACK_MST_BNDL_WRP_7340 = None
FILLER_44 = None
FILLER_62 = None

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
    packing_machine: str = 'MACH5'  # Packing machine used for this bundle
    skus: List[PlacedSKU] = None

    def __init__(self, width: float, height: float, max_length: float = 3680, packing_machine: str = 'MACH5'):
        self.width = width
        self.height = height
        self.max_length = max_length
        self.packing_machine = packing_machine
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

    def get_actual_dimensions(self, visual=False):
        """
        Calculate the actual dimensions of the bundle based on placed SKUs
        """
        if not self.skus:
            return 0, 0, 0
        if visual:
            non_packaging_skus = [sku for sku in self.skus if not (sku.id.startswith("Pack_") and "Filler" not in sku.id)]
        else:
            non_packaging_skus = [sku for sku in self.skus if not sku.id.startswith("Pack_")]

        max_x = max(sku.x + sku.width for sku in non_packaging_skus)
        max_y = max(sku.y + sku.height for sku in non_packaging_skus)
        # max_length = max(sku.length for sku in self.skus)
        max_length = 3680 if (max(sku.length for sku in non_packaging_skus if sku.length) < 3700) else 7340

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
            self.add_sku(PACK_SUB_BNDL_WRP_3680, 0, 0, False)
            self.add_sku(PACK_MST_BNDL_WRP_3680, 0, 0, False)

            self.max_length = 3680

            # Pack pad weight
            if width >= 152: # If less than 6 inches, don't add any pads
                if width <= 216: # 8.5 inches
                    self.add_sku(PACK_PAD_8_3680, 0, 0, False)
                elif width <= 254: # 10 inches
                    self.add_sku(PACK_PAD_10_3680, 0, 0, False)
                elif width <= 331: # 13 inches
                    self.add_sku(PACK_PAD_13_3680, 0, 0, False)
                else: # 19 inches
                    self.add_sku(PACK_PAD_19_3680, 0, 0, False)

            if height >= 152: # If less than 6 inches, don't add any pads
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
            self.add_sku(PACK_SUB_BNDL_WRP_7340, 0, 0, False)
            self.add_sku(PACK_MST_BNDL_WRP_7340, 0, 0, False)

            self.max_length = 7340

            # Pack pad weight
            if width >= 152: # If less than 6 inches, don't add any pads
                if width <= 216: # 8.5 inches
                    self.add_sku(PACK_PAD_8_7340, 0, 0, False)
                elif width <= 254: # 10 inches
                    self.add_sku(PACK_PAD_10_7340, 0, 0, False)
                elif width <= 331: # 13 inches
                    self.add_sku(PACK_PAD_13_7340, 0, 0, False)
                else: # 19 inches
                    self.add_sku(PACK_PAD_19_7340, 0, 0, False)

            if height >= 152: # If less than 6 inches, don't add any pads
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

def create_packaging_classes(data: List[dict]) -> List[SKU]:
    """
    Create SKU classes for packaging and filler materials from the provided data.
    """
    global PACK_ANGLE_3680, PACK_ANGLE_7340, PACK_1_4_19_DUN_3680, PACK_1_4_19_DUN_7340
    global PACK_2_3_19_DUN_3680, PACK_2_3_19_DUN_7340, PACK_LUMBER_3680, PACK_LUMBER_7340
    global PACK_PAD_8_3680, PACK_PAD_8_7340, PACK_PAD_10_3680, PACK_PAD_10_7340
    global PACK_PAD_13_3680, PACK_PAD_13_7340, PACK_PAD_19_3680, PACK_PAD_19_7340
    global PACK_SUB_BNDL_WRP_3680, PACK_SUB_BNDL_WRP_7340
    global PACK_MST_BNDL_WRP_3680, PACK_MST_BNDL_WRP_7340
    global FILLER_44, FILLER_62

    # Pack_Angle
    pack_angle = data['Pack_Angle']
    PACK_ANGLE_3680 = SKU(
        id='Pack_Angle_3680',
        bundleqty=pack_angle['3680mm Qty'],
        width=pack_angle['Width (mm)'],
        height=pack_angle['Height (mm)'],
        length=pack_angle['3680mm Length (mm)'],
        weight=pack_angle['3680mm Weight (kg)'],
        desc=pack_angle['Description'],
    )

    PACK_ANGLE_7340 = SKU(
        id='Pack_Angle_7340',
        bundleqty=pack_angle['7340mm Qty'],
        width=pack_angle['Width (mm)'],
        height=pack_angle['Height (mm)'],
        length=pack_angle['7340mm Length (mm)'],
        weight=pack_angle['7340mm Weight (kg)'],
        desc=pack_angle['Description'],
    )

    # Pack_1x4x19_Dun
    pack_1x4x19_dun = data['Pack_1x4x19_Dun']
    PACK_1_4_19_DUN_3680 = SKU(
        id='Pack_1x4x19_Dun_3680',
        bundleqty=pack_1x4x19_dun['3680mm Qty'],
        width=pack_1x4x19_dun['Width (mm)'],
        height=pack_1x4x19_dun['Height (mm)'],
        length=pack_1x4x19_dun['3680mm Length (mm)'],
        weight=pack_1x4x19_dun['3680mm Weight (kg)'],
        desc=pack_1x4x19_dun['Description'],
    )

    PACK_1_4_19_DUN_7340 = SKU(
        id='Pack_1x4x19_Dun_7340',
        bundleqty=pack_1x4x19_dun['7340mm Qty'],
        width=pack_1x4x19_dun['Width (mm)'],
        height=pack_1x4x19_dun['Height (mm)'],
        length=pack_1x4x19_dun['7340mm Length (mm)'],
        weight=pack_1x4x19_dun['7340mm Weight (kg)'],
        desc=pack_1x4x19_dun['Description'],
    )

    # Pack_2x3x19_Dun
    pack_2x3x19_dun = data['Pack_2x3x19_Dun']
    PACK_2_3_19_DUN_3680 = SKU(
        id='Pack_2x3x19_Dun_3680',
        bundleqty=pack_2x3x19_dun['3680mm Qty'],
        width=pack_2x3x19_dun['Width (mm)'],
        height=pack_2x3x19_dun['Height (mm)'],
        length=pack_2x3x19_dun['3680mm Length (mm)'],
        weight=pack_2x3x19_dun['3680mm Weight (kg)'],
        desc=pack_2x3x19_dun['Description'],
    )
    PACK_2_3_19_DUN_7340 = SKU(
        id='Pack_2x3x19_Dun_7340',
        bundleqty=pack_2x3x19_dun['7340mm Qty'],
        width=pack_2x3x19_dun['Width (mm)'],
        height=pack_2x3x19_dun['Height (mm)'],
        length=pack_2x3x19_dun['7340mm Length (mm)'],
        weight=pack_2x3x19_dun['7340mm Weight (kg)'],
        desc=pack_2x3x19_dun['Description'],
    )

    # Pack_Lumber
    pack_lumber = data['Pack_Lumber']
    PACK_LUMBER_3680 = SKU(
        id='Pack_Lumber_3680',
        bundleqty=pack_lumber['3680mm Qty'],
        width=pack_lumber['Width (mm)'],
        height=pack_lumber['Height (mm)'],
        length=pack_lumber['3680mm Length (mm)'],
        weight=pack_lumber['3680mm Weight (kg)'],
        desc=pack_lumber['Description'],
    )
    PACK_LUMBER_7340 = SKU(
        id='Pack_Lumber_7340',
        bundleqty=pack_lumber['7340mm Qty'],
        width=pack_lumber['Width (mm)'],
        height=pack_lumber['Height (mm)'],
        length=pack_lumber['7340mm Length (mm)'],
        weight=pack_lumber['7340mm Weight (kg)'],
        desc=pack_lumber['Description'],
    )

    # Pack_Pad_8
    pack_pad_8 = data['Pack_Pad_8']
    PACK_PAD_8_3680 = SKU(
        id='Pack_Pad_8_3680',
        bundleqty=pack_pad_8['3680mm Qty'],
        width=pack_pad_8['Width (mm)'],
        height=pack_pad_8['Height (mm)'],
        length=pack_pad_8['3680mm Length (mm)'],
        weight=pack_pad_8['3680mm Weight (kg)'],
        desc=pack_pad_8['Description'],
    )
    PACK_PAD_8_7340 = SKU(
        id='Pack_Pad_8_7340',
        bundleqty=pack_pad_8['7340mm Qty'],
        width=pack_pad_8['Width (mm)'],
        height=pack_pad_8['Height (mm)'],
        length=pack_pad_8['7340mm Length (mm)'],
        weight=pack_pad_8['7340mm Weight (kg)'],
        desc=pack_pad_8['Description'],
    )

    # Pack_Pad_10
    pack_pad_10 = data['Pack_Pad_10']
    PACK_PAD_10_3680 = SKU(
        id='Pack_Pad_10_3680',
        bundleqty=pack_pad_10['3680mm Qty'],
        width=pack_pad_10['Width (mm)'],
        height=pack_pad_10['Height (mm)'],
        length=pack_pad_10['3680mm Length (mm)'],
        weight=pack_pad_10['3680mm Weight (kg)'],
        desc=pack_pad_10['Description'],
    )
    PACK_PAD_10_7340 = SKU(
        id='Pack_Pad_10_7340',
        bundleqty=pack_pad_10['7340mm Qty'],
        width=pack_pad_10['Width (mm)'],
        height=pack_pad_10['Height (mm)'],
        length=pack_pad_10['7340mm Length (mm)'],
        weight=pack_pad_10['7340mm Weight (kg)'],
        desc=pack_pad_10['Description'],
    )

    # Pack_Pad_13
    pack_pad_13 = data['Pack_Pad_13']
    PACK_PAD_13_3680 = SKU(
        id='Pack_Pad_13_3680',
        bundleqty=pack_pad_13['3680mm Qty'],
        width=pack_pad_13['Width (mm)'],
        height=pack_pad_13['Height (mm)'],
        length=pack_pad_13['3680mm Length (mm)'],
        weight=pack_pad_13['3680mm Weight (kg)'],
        desc=pack_pad_13['Description'],
    )
    PACK_PAD_13_7340 = SKU(
        id='Pack_Pad_13_7340',
        bundleqty=pack_pad_13['7340mm Qty'],
        width=pack_pad_13['Width (mm)'],
        height=pack_pad_13['Height (mm)'],
        length=pack_pad_13['7340mm Length (mm)'],
        weight=pack_pad_13['7340mm Weight (kg)'],
        desc=pack_pad_13['Description'],
    )

    # Pack_Pad_19
    pack_pad_19 = data['Pack_Pad_19']
    PACK_PAD_19_3680 = SKU(
        id='Pack_Pad_19_3680',
        bundleqty=pack_pad_19['3680mm Qty'],
        width=pack_pad_19['Width (mm)'],
        height=pack_pad_19['Height (mm)'],
        length=pack_pad_19['3680mm Length (mm)'],
        weight=pack_pad_19['3680mm Weight (kg)'],
        desc=pack_pad_19['Description'],
    )
    PACK_PAD_19_7340 = SKU(
        id='Pack_Pad_19_7340',
        bundleqty=pack_pad_19['7340mm Qty'],
        width=pack_pad_19['Width (mm)'],
        height=pack_pad_19['Height (mm)'],
        length=pack_pad_19['7340mm Length (mm)'],
        weight=pack_pad_19['7340mm Weight (kg)'],
        desc=pack_pad_19['Description'],
    )

    # Pack_Sub_Bndl_Wrp
    pack_sub_bndl_wrp = data['Pack_Sub_Bndl_Wrp']
    PACK_SUB_BNDL_WRP_3680 = SKU(
        id='Pack_Sub_Bndl_Wrp_3680',
        bundleqty=pack_sub_bndl_wrp['3680mm Qty'],
        width=pack_sub_bndl_wrp['Width (mm)'],
        height=pack_sub_bndl_wrp['Height (mm)'],
        length=pack_sub_bndl_wrp['3680mm Length (mm)'],
        weight=pack_sub_bndl_wrp['3680mm Weight (kg)'],
        desc=pack_sub_bndl_wrp['Description'],
    )
    PACK_SUB_BNDL_WRP_7340 = SKU(
        id='Pack_Sub_Bndl_Wrp_7340',
        bundleqty=pack_sub_bndl_wrp['7340mm Qty'],
        width=pack_sub_bndl_wrp['Width (mm)'],
        height=pack_sub_bndl_wrp['Height (mm)'],
        length=pack_sub_bndl_wrp['7340mm Length (mm)'],
        weight=pack_sub_bndl_wrp['7340mm Weight (kg)'],
        desc=pack_sub_bndl_wrp['Description'],
    )

    # Pack_Mst_Bndl_Wrp
    pack_mst_bndl_wrp = data['Pack_Mst_Bndl_Wrp']
    PACK_MST_BNDL_WRP_3680 = SKU(
        id='Pack_Mst_Bndl_Wrp_3680',
        bundleqty=pack_mst_bndl_wrp['3680mm Qty'],
        width=pack_mst_bndl_wrp['Width (mm)'],
        height=pack_mst_bndl_wrp['Height (mm)'],
        length=pack_mst_bndl_wrp['3680mm Length (mm)'],
        weight=pack_mst_bndl_wrp['3680mm Weight (kg)'],
        desc=pack_mst_bndl_wrp['Description'],
    )
    PACK_MST_BNDL_WRP_7340 = SKU(
        id='Pack_Mst_Bndl_Wrp_7340',
        bundleqty=pack_mst_bndl_wrp['7340mm Qty'],
        width=pack_mst_bndl_wrp['Width (mm)'],
        height=pack_mst_bndl_wrp['Height (mm)'],
        length=pack_mst_bndl_wrp['7340mm Length (mm)'],
        weight=pack_mst_bndl_wrp['7340mm Weight (kg)'],
        desc=pack_mst_bndl_wrp['Description'],
    )

    # Filler materials
    filler_44 = data['Pack_44Filler']
    FILLER_44 = SKU(
        id='Pack_44Filler',
        bundleqty=filler_44['3680mm Qty'],
        width=filler_44['Width (mm)'],
        height=filler_44['Height (mm)'],
        length=filler_44['3680mm Length (mm)'],
        weight=filler_44['3680mm Weight (kg)'],
        desc=filler_44['Description'],
    )
    filler_62 = data['Pack_62Filler']
    FILLER_62 = SKU(
        id='Pack_62Filler',
        bundleqty=filler_62['3680mm Qty'],
        width=filler_62['Width (mm)'],
        height=filler_62['Height (mm)'],
        length=filler_62['3680mm Length (mm)'],
        weight=filler_62['3680mm Weight (kg)'],
        desc=filler_62['Description'],
    )

    # Get packaging heights
    packaging_height = (2*pack_angle['Height (mm)'] +
                        pack_1x4x19_dun['Height (mm)'] +
                        pack_2x3x19_dun['Height (mm)'] +
                        2*pack_sub_bndl_wrp['Height (mm)'])
    # Use height measurements for width too,
    # as it is still height but just adds to the side of the bundle
    packaging_width = (2*pack_angle['Height (mm)'] +
                       2*pack_sub_bndl_wrp['Height (mm)'])
    lumber_height = (pack_lumber['Height (mm)'])
    return packaging_height, packaging_width, lumber_height