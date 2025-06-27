import openpyxl
import pandas as pd
from tkinter import filedialog
from typing import List, Tuple, Optional

from bundle_classes import SKU, Bundle
from bundle_visualize import visualize_bundles

global maxWidth
global maxHeight
global workingDir

maxWidth = 590
maxHeight = 590

def get_workbook():
    """
    Open a dialog to pick an Excel file, return the workbook object
    """
    global workingDir
    # get excel file path from user
    path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx; *.xls; *.xlsm")],
    )
    if not path:
        print("No workbook selected.")
        return None

    workingDir = path.rsplit('/', 1)[0]  # get the directory of the selected file

    # load the workbook and read the sheet "SO_Input"
    workbook = openpyxl.load_workbook(path)
    return workbook

def get_data(workbook):
    """
    Read data from the 'SO_Input' sheet of the workbook
    """
    # get the "SO_Input" sheet
    if "SO_Input" not in workbook.sheetnames:
        print("Sheet 'SO_Input' not found in the selected file.")
        return None
    sheet = workbook["SO_Input"]

    # create headers with dataframe
    headers = [cell.value for cell in sheet[1]] # first row as headers
    df = pd.DataFrame(columns=headers)
    # read all rows from the sheet
    for row in sheet.iter_rows(min_row=2, values_only=True):
        df.loc[len(df)] = row  # append each row to the dataframe

    return df

def pack_skus(skus: List[SKU], bundle_width: int, bundle_height: int) -> List[Bundle]:
    skus_sorted = sorted(skus, key=lambda s: -s.weight)
    bundles: List[Bundle] = []

    for sku in skus_sorted:
        placed = False
        for bundle in bundles:
            if bundle.try_place_sku(sku):
                placed = True
                break
        if not placed:
            new_bundle = Bundle(bundle_width, bundle_height)
            new_bundle.try_place_sku(sku)
            bundles.append(new_bundle)

    # Add filler materials to each bundle after main SKUs are packed
    for bundle in bundles:
        bundle.add_filler_materials()

    return bundles

def create_sku_objects(order_rows: dict):
    """
    Create arrays of SKU objects for each order
    """
    order_skus = {}
    for order, rows in order_rows.items():
        skus = []
        for _, row in rows.iterrows():
            # get quantity of SKU from the row
            quantity = row['SKU Qty']
            for _ in range(quantity):
                sku = SKU(
                    id=row['SKU'],
                    width=row['SKU Width (mm)'],
                    height=row['SKU Height (mm)'],
                    length=row['SKU Length (mm)'],
                    weight=row['SKU Weight (kg)'],
                    desc=row['SKU Description']
                )
                skus.append(sku)
        order_skus[order] = skus
    return order_skus

def write_optimized_bundles(workbook, order_bundles: dict):
    """
    Write the packed bundles to a new sheet in the workbook
    """
    # create a new sheet for the optimized bundles
    if "Optimized_Bundles" in workbook.sheetnames:
        del workbook["Optimized_Bundles"]
    optimized_sheet = workbook.create_sheet("Optimized_Bundles")

    # write headers
    headers = [
               "Order ID",
               "Bundle No.",
               "SKU",
               "Qty",
               "SKU Description",
            #    "SKU Width (mm)",
            #    "SKU Height (mm)",
            #    "SKU Weight (kg)",
               "Bundle Total Width (mm)",
               "Bundle Total Height (mm)",
               "Bundle Total Weight (kg)",
               "Note"
               ]
    optimized_sheet.append(headers)

    # add data for each order's bundles
    for order, bundles in order_bundles.items():
        for bundle_index, bundle in enumerate(bundles):
            # get quantity of each SKU in the bundle
            sku_counts = {}
            for sku in bundle.skus:
                if sku.id not in sku_counts:
                    sku_counts[sku.id] = {'qty': 0, 'description': sku.desc, 'width': sku.width, 'height': sku.height, 'weight': sku.weight}
                sku_counts[sku.id]['qty'] += 1
            # calculate bundle total width, height, and weight
            total_width = bundle.width
            total_height = bundle.height
            total_weight = sum(sku.weight for sku in bundle.skus)
            # write each SKU in the bundle to the sheet
            for sku_id, sku_data in sku_counts.items():
                optimized_sheet.append([
                    order,
                    bundle_index + 1,
                    sku_id,
                    sku_data['qty'],
                    sku_data['description'],
                    # sku_data['width'],
                    # sku_data['height'],
                    # sku_data['weight'],
                    total_width,
                    total_height,
                    total_weight,
                    ""
                ])
        # add a blank row after each order's bundles
        optimized_sheet.append([])

    # create a table over the data
    table = openpyxl.worksheet.table.Table(displayName="OptimizedBundlesTable", ref=optimized_sheet.dimensions, tableStyleInfo=openpyxl.worksheet.table.TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True))
    # resize the table to fit the data
    optimized_sheet.add_table(table)

    # save the workbook
    try:
        workbook.save(f"{workingDir}/Optimized_Bundles.xlsx")
    except Exception as e:
        print("Error saving the file. Is it already open?")
        return

if __name__ == "__main__":
    # get the workbook and data
    print("Please select the Excel file with order data")
    wb = get_workbook()
    if not wb:
        exit()
    data = get_data(wb)

    print("\nData loaded successfully! Sorting and packing SKUs...\n")

    # get unique orders
    unique_orders = data['Order ID'].unique()

    # get array of rows in sorted_data that belong to each order
    order_rows = {order: data[data['Order ID'] == order] for order in unique_orders}

    # create SKU objects for each order
    order_skus = create_sku_objects(order_rows)

    # ask user about saving figures
    saveFigs = input("Do you want to save packing visualizations as images? (y/n): ").strip().lower() == 'y'
    # create a directory for images if needed
    if saveFigs:
        print("Saving images...")
        import os
        images_dir = f"{workingDir}/images"
        os.makedirs(images_dir, exist_ok=True)

    # pack each order's SKUs into bundles
    order_bundles = {}
    for order, skus in order_skus.items():
        bundles = pack_skus(skus, maxWidth, maxHeight)
        order_bundles[order] = bundles
        for i, bundle in enumerate(bundles):
            for sku in bundle.skus:
                rot = " (rotated)" if sku.rotated else ""
        if saveFigs:
            visualize_bundles(bundles, f"{images_dir}/Order_{order}.png")

    if saveFigs:
        print(f"\nImages saved in {images_dir}")

    # write the packed bundles to a new sheet in the workbook
    write_optimized_bundles(wb, order_bundles)

    print(f"Optimized bundles written to {workingDir}/Optimized_Bundles.xlsx\n\nProgram complete.")