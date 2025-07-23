from PyQt6.QtWidgets import QMessageBox, QWidget
from BundleQtGui import Ui_BundleOptimizer

import openpyxl
import pandas as pd
from tkinter import filedialog
import os
from math import ceil, floor
from datetime import datetime
import numpy as np

from bundle_classes import SKU
from bundle_visualize import visualize_bundles
from bundle_packing import pack_skus

class Ui_MainWindow:
    def __init__(self):
        self.Widget = QWidget()
        self.ui = Ui_BundleOptimizer()
        self.ui.setupUi(self.Widget)

        self.setupUi()
        self.Widget.show()

    def setupUi(self):
        # Initialize values and connect signals (camelCase, while helper methods are snake_case)
        self.ui.fileBrowse.clicked.connect(self.getInputWorkbook)
        self.ui.appendBrowse.clicked.connect(self.getAppendWorkbook)
        self.ui.openExample.clicked.connect(self.openExampleFile)
        self.ui.optimizeBundles.clicked.connect(self.optimizeBundles)
        self.ui.openImages.clicked.connect(self.openImages)
        self.ui.openExcel.clicked.connect(self.openExcel)
        self.ui.helpButton.clicked.connect(self.openHelp)

        self.workingDir = None  # to hold the directory of the selected Excel file

## QT Connection Methods (camelCase)

    def getInputWorkbook(self):
        """
        Open a dialog to pick an Excel file, return the workbook object
        """
        # get excel file path from user
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx; *.xls; *.xlsm")],
        )
        if not path:
            return

        self.workingDir = path.rsplit('/', 1)[0]  # get the directory of the selected file

        # load the workbook and read the sheet "SO_Input"
        self.ui.excelDir.setText(path)

    def getAppendWorkbook(self):
        """
        Get the workbook of Optimized data to append new data to
        """
        path = filedialog.askopenfilename(
            title="Select an Excel file to append optimized data to",
            filetypes=[("Excel files", "*.xlsx; *.xls; *.xlsm")],
        )
        if not path:
            return

        # load the workbook and read the sheet "SO_Input"
        self.ui.appendDir.setText(path)

    def openExampleFile(self):
        """
        Open the example Excel file to the user, in Excel
        """
        example_path = os.path.join(os.path.dirname(__file__), 'SO_Input_Example.xlsx')
        if os.path.exists(example_path):
            os.startfile(example_path)
        else:
            self.show_alert("File Not Found", "Example file not found.")

    def optimizeBundles(self):
        """
        Optimize bundles based on the data from the Excel file
        """
        self.maxWidth = 559
        self.maxHeight = 559
        self.maxLength = 3680
        self.missingDataSKUs = []  # to hold SKUs that are missing data in the Excel file
        self.append_data = False  # to indicate if we are appending data to an existing workbook

        self.ui.progressLabel.setText("Getting data...")
        self.ui.progressBar.setValue(10)
        # Get data from input workbook
        try:
            self.workbook = openpyxl.load_workbook(self.ui.excelDir.text())
            data = self.get_data(self.workbook)
        except Exception as e:
            self.show_alert("Error", "Invalid path for Excel file.", "error")
            self.ui.progressBar.setValue(0)
            self.ui.progressLabel.setText("")
            return

        # get unique orders
        unique_orders = list(data['OrderNbr'].unique())

        if self.ui.appendDir.text():
            self.appendWorkbook = openpyxl.load_workbook(self.ui.appendDir.text())
            self.append_data = True
            unique_orders, self.appendWorkbook = self.remove_optimized_orders(unique_orders, self.appendWorkbook)
            if not unique_orders:
                self.show_alert("No Orders", "All selected orders have already been optimized in the append workbook.", "info")
                self.ui.progressBar.setValue(0)
                self.ui.progressLabel.setText("")
                return
        else:
            # delete existing file
            self.append_data = False
            try:
                if os.path.exists(f"{self.workingDir}/Optimized_Bundles.xlsx"):
                    os.remove(f"{self.workingDir}/Optimized_Bundles.xlsx")
            except Exception as e:
                self.show_alert("Error", f"Error accessing Optimized_Bundles.xlsx file: {e}", "error")
                self.ui.progressBar.setValue(0)
                self.ui.progressLabel.setText("")
                return

        # get array of rows in sorted_data that belong to each order
        order_rows = {order: data[data['OrderNbr'] == order] for order in unique_orders}

        # create SKU objects for each order
        order_skus = self.create_sku_objects(order_rows)

        order_skus = self.remove_invalids(order_skus)

        self.ui.progressBar.setValue(20)
        images_dir = f"{self.workingDir}/images"
        os.makedirs(images_dir, exist_ok=True)

        images_dir = f"{self.workingDir}/images"
        os.makedirs(images_dir, exist_ok=True)

        # pack each order's SKUs into bundles
        order_bundles = {}
        for order, skus in order_skus.items():
            if skus == []:
                continue
            self.ui.progressBar.setValue(int(round(20 + 70 * (list(order_skus.keys()).index(order) + 1) / len(order_skus))))
            self.ui.progressLabel.setText(f"Packing order {order}...")
            bundles, self.removed_skus = pack_skus(skus, self.maxWidth, self.maxHeight)
            order_bundles[order] = bundles

            visualize_bundles(bundles, f"{images_dir}/Order_{order}.png")
        self.images_dir = images_dir

        # write the packed bundles to a new sheet in the workbook
        if self.append_data:
            self.workbook = self.appendWorkbook
        self.ui.progressBar.setValue(90)
        self.ui.progressLabel.setText("Writing optimized bundles to Excel...")
        self.missingDataSKUs.extend(self.removed_skus)  # add removed SKUs to missing data SKUs
        self.write_optimized_bundles(self.workbook, order_bundles)

        self.ui.progressBar.setValue(100)
        self.ui.progressLabel.setText("Packing complete!")

        # if self.missingDataSKUs:
        #     self.show_alert("Missing Data", "There exist InventoryIDs that are missing data in the Excel file\nand have been excluded from optimization.\n\nPlease check the 'missing_data_skus.txt' file for details.", "warning")
        #     # write the missing SKUs to a file
        #     with open(f"{self.workingDir}/missing_data_skus.txt", 'w') as f:
        #         f.write("The following SKUs are missing data and could not be included in the optimization:\n\n")
        #         # group by order number
        #         missing_skus_by_order = {}
        #         for sku in self.missingDataSKUs:
        #             order_nbr = sku.data['OrderNbr']
        #             if order_nbr not in missing_skus_by_order:
        #                 missing_skus_by_order[order_nbr] = []
        #             if sku.id not in missing_skus_by_order[order_nbr]:
        #                 missing_skus_by_order[order_nbr].append(sku.id)
        #         for order_nbr, skus in missing_skus_by_order.items():
        #             f.write(f"Order {order_nbr}:\n")
        #             for sku in skus:
        #                 f.write(f"- {sku}\n")
        #             f.write("\n")
        # if self.removed_skus:
        #     self.show_alert("Removed SKUs", "There exist InventoryIDs that were removed from the optimization process due to their dimensions.\n\nPlease check the 'removed_skus.txt' file for details.", "warning")
        #     # group removed SKUs by order number
        #     removed_skus_by_order = {}
        #     for sku in self.removed_skus:
        #         order_nbr = sku.data['OrderNbr']
        #         if order_nbr not in removed_skus_by_order:
        #             removed_skus_by_order[order_nbr] = []
        #         if sku.id not in [sku.id for sku in removed_skus_by_order[order_nbr]]:
        #             removed_skus_by_order[order_nbr].append(sku)
        #     # write the removed SKUs to a file
        #     with open(f"{self.workingDir}/removed_skus.txt", 'w') as f:
        #         f.write("The following SKUs were removed from the optimization process due to their dimensions:\n\n")
        #         for order_nbr, skus in removed_skus_by_order.items():
        #             f.write(f"Order {order_nbr}:\n")
        #             for sku in skus:
        #                 f.write(f"- {sku.id} (Width: {sku.width}, Height: {sku.height}, Length: {sku.length}, Weight: {sku.weight})\n")
        #             f.write("\n")

    def openImages(self):
        """
        Open the images directory in the file explorer
        """
        if hasattr(self, 'images_dir'):
            os.startfile(self.images_dir)
        else:
            self.show_alert("No Images", "No images have been generated yet. Please optimize bundles first.", "error")

    def openExcel(self):
        """
        Open the optimized bundles Excel file in the file explorer
        """
        if not self.workingDir:
            self.show_alert("Error", "No file found", "error")
            return
        excel_path = f"{self.workingDir}/Optimized_Bundles.xlsx"
        if os.path.exists(excel_path):
            os.startfile(excel_path)
        else:
            self.show_alert("Error", "Optimized bundles file not found. Please optimize bundles first.", "error")

    def openHelp(self):
        """
        Open the help file in the file explorer
        """
        self.show_alert("Help", """
        For additional support, please contact Aionex Solutions Ltd.
        (604-309-8975)

        To use this tool, follow these steps:

        1. Click on the first 'Browse' button to select your Excel file
           containing SKU data.
            - If you are unsure how to format your Excel file,
              you can click on 'Open Example File' to open a sample file.
            - Please format your Excel file according to the example provided.

        2. If you want to append the optimized data to an existing file,
           click on the second 'Browse' button to select the Excel file
           where you want to append the optimized data.

        3. Click on 'Perform Bundle Optimization' to start the
           optimization process.
           Wait for the progress bar to complete.

        4. Once the optimization is complete, you can:
            - Click on 'Open Images Folder' to view the generated bundle images.
            - Click on 'Open Resultant Excel File' to view the optimized
              bundle data in an Excel file.

        NOTE: All files are generated in the same directory as the input Excel file.
        """, type="info")

## Helper Methods (snake_case)

    def get_sub_bundle_data_sheet(self):
        """
        Open a dialog to pick an Excel file, return the sheet "Sub-Bundle_Data"
        """
        # load the workbook and read the sheet "Sub-Bundle_Data"
        path = os.path.join(os.path.dirname(__file__), 'Sub-Bundle_Data.xlsx')
        workbook = openpyxl.load_workbook(path)
        sheet = workbook["Sub-Bundle_Data"]
        if not sheet:
            self.show_alert("Warning", "Sheet 'Sub-Bundle_Data' not found in file.")
            return None
        return sheet

    def remove_optimized_orders(self, orders, workbook):
        """
        Remove orders that are already optimized in the workbook
        """
        if "Optimized_Bundles" not in workbook.sheetnames:
            return orders, workbook
        optimized_sheet = workbook["Optimized_Bundles"]
        # remove table formatting if it exists
        if "OptimizedBundlesTable" in optimized_sheet.tables:
            del optimized_sheet.tables["OptimizedBundlesTable"]
        # find orders that have been updated since last optimization (the 'OptimizedOn' column is earlier than the 'LastModifiedOn' column)
        found_order = False
        for order in orders:
            # get the last modified date of the order from the optimized sheet
            for rowIdx in range(optimized_sheet.max_row, 0, -1):
                row = optimized_sheet[rowIdx]
                if row[1].value == order:
                    last_modified_on = row[-2].value
                    optimized_on = row[-1].value
                    # convert to datetime if not None
                    if last_modified_on and optimized_on:
                        last_modified_on = datetime.strptime(last_modified_on, "%Y-%m-%d")
                        optimized_on = datetime.strptime(optimized_on, "%Y-%m-%d")
                        if last_modified_on <= optimized_on: # modified earlier than it was optimized
                            orders = [o for o in orders if o != order]  # remove the order from the list
                        else:
                            if not found_order:
                                found_order = True
                                optimized_sheet.row_dimensions[rowIdx + 1].outlineLevel = 0
                                optimized_sheet.delete_rows(rowIdx + 1, 1)  # remove blank row that separates orders
                            optimized_sheet.row_dimensions[rowIdx].outlineLevel = 0
                            optimized_sheet.delete_rows(rowIdx, 1)  # remove the row from the optimized sheet
            found_order = False

        # save the workbook after removing orders
        try:
            optimized_sheet.append([])  # add a blank row at the end to separate orders
            workbook.save(self.ui.appendDir.text())
        except Exception as e:
            self.show_alert("Error", f"Error with the existing optimized data file. Is it already open? Error: {e}", "error")
            return [], workbook
        return orders, workbook

    def get_data(self, workbook):
        """
        Read data from the 'SO_Input' sheet of the workbook
        """
        # get the "SO_PackExportData" sheet
        so_input = workbook["SO-PackExportData"]
        if not so_input:
            self.show_alert("Warning", "Sheet 'SO-PackExportData' is empty or not found. Using first sheet in the file instead.")
            so_input = workbook.active  # fallback to the first sheet if is not found
        sb_data = self.get_sub_bundle_data_sheet()

        df = pd.DataFrame(columns=[cell.value for cell in so_input[1]])
        sb_df = pd.DataFrame(columns=[cell.value for cell in sb_data[2]])
        # read all rows from the sheets
        for row in so_input.iter_rows(min_row=2, values_only=True):
            if all(cell is None for cell in row):
                continue
            df.loc[len(df)] = row
        for row in sb_data.iter_rows(min_row=3, values_only=True):
            # check for any empty cells in the row
            if any(cell is None for cell in [el for el in row][2:7]):
                continue
            sb_df.loc[len(sb_df)] = row

        # check if the required columns are present
        self.headers = ["OrderType", "OrderNbr", "Bdl_Override", "InventoryID", "Quantity", "Pcs/Bundle", "Can_be_bottom",
                   "Dim_shrink", "Width_mm", "Height_mm", "Length_mm", "Weight_kg", "UOM", "Description",
                   "ShipTo", "AddressLine1", "AddressLine2", "City", "State", "Country", "Status", "OrderDate",
                   "ProdReleaseDate", "SchedShipDate", "TargetArrival", "NotBefore", "ShipVia", "LastModifiedOn"]
        for col in self.headers:
            if col not in df.columns:
                # add the column with default values
                df[col] = None

        # put data from sb_df into df rows
        for _, row in sb_df.iterrows():
            sku_id = row['SKU']
            df_rows = []
            if row['Bottom Row Acceptable']:
                can_be_bottom = True
            else:
                can_be_bottom = False
            for index, df_row in df.iterrows():
                try:
                    if sku_id in df_row['InventoryID']:
                        df_rows.append(index)
                except TypeError:
                    # if df_row['InventoryID'] is None, skip this row
                    continue
            if row['Partial Dim To Reduce'] is None:
                row['Partial Dim To Reduce'] = ''
            for df_row_idx in df_rows:
                # update the Pcs/Bundle column with the value from sb_df
                df.loc[df_row_idx, 'Pcs/Bundle'] = row['Qty/bundle']
                df.loc[df_row_idx, 'Width_mm'] = row['Width (mm)']
                df.loc[df_row_idx, 'Height_mm'] = row['Height (mm)']
                df.loc[df_row_idx, 'Length_mm'] = row['Length (mm)']
                df.loc[df_row_idx, 'Weight_kg'] = row['Weight kg/length']
                df.loc[df_row_idx, 'Dim_shrink'] = row['Partial Dim To Reduce']
                df.loc[df_row_idx, 'Can_be_bottom'] = can_be_bottom

        # iterate through df and convert qty (in pieces) to qty (in bundles)
        for index, row in df.iterrows():
            if row['Pcs/Bundle'] is not None and row['Quantity'] is not None:
                try:
                    # convert Quantity to Pcs/Bundle
                    whole_qty = floor(ceil(abs(row['Quantity'])) / row['Pcs/Bundle'])
                    fraction_remaining = (ceil(abs(row['Quantity'])) % row['Pcs/Bundle']) / row['Pcs/Bundle']
                    if fraction_remaining > 0:
                        df.loc[index, 'Quantity'] = float(whole_qty) + float(fraction_remaining)
                    else:
                        df.loc[index, 'Quantity'] = int(whole_qty)
                except ZeroDivisionError:
                    self.show_alert("Error", f"Pcs/Bundle cannot be zero for SKU {row['InventoryID']}. Please check the input data.", "error")
                    return pd.DataFrame()

        return df

    def remove_invalids(self, order_skus: dict) -> dict:
        """
        Iterate through the order_skus dictionary and remove any SKUs that have None as width, height, length, or weight.
        """
        for order, skus in order_skus.items():
            valid_skus = []
            for sku in skus:
                if None in (sku.width, sku.height, sku.length, sku.weight):
                    if sku.id not in self.missingDataSKUs:
                        self.missingDataSKUs.append(sku)  # add to missing data SKUs list
                else:
                    valid_skus.append(sku)
            order_skus[order] = valid_skus
        return order_skus

    def create_sku_objects(self, order_rows: dict):
        """
        Create arrays of SKU objects for each order
        """
        order_skus = {}
        for order, rows in order_rows.items():
            skus = []
            for _, row in rows.iterrows():
                # get quantity of SKU from the row
                quantity = row['Quantity']
                invID = row['InventoryID'].strip()
                newLength = row['Length_mm']
                if newLength is not None:
                    if 3600 <= newLength <= 3700:
                        newLength = 3650

                if type(quantity) is float:
                    # partial sub-bundle
                    remainder = quantity - floor(quantity)
                    if remainder > 0:
                        if not row['Weight_kg']:
                            continue
                        weight = row['Weight_kg'] * quantity
                        width, height = self.shrink_to_square(row['Width_mm'], row['Height_mm'], remainder, row['Dim_shrink'])
                        new_invID = f"{invID}_Partial"
                        sku = SKU(
                            id=new_invID,
                            bundleqty=row['Pcs/Bundle'] * remainder,
                            width=width,
                            height=height,
                            length=newLength,
                            weight=weight,
                            desc=row['Description'],
                            can_be_bottom=row['Can_be_bottom'],
                            data={
                                'OrderType': row['OrderType'],
                                'OrderNbr': row['OrderNbr'],
                                'UOM': row['UOM'],
                                'Bdl_Override': row['Bdl_Override'] if pd.notna(row['Bdl_Override']) else None,
                                'ShipTo': row['ShipTo'],
                                'AddressLine1': row['AddressLine1'],
                                'AddressLine2': row['AddressLine2'],
                                'City': row['City'],
                                'State': row['State'],
                                'Country': row['Country'],
                                'Status': row['Status'],
                                'OrderDate': row['OrderDate'],
                                'ProdReleaseDate': row['ProdReleaseDate'],
                                'SchedShipDate': row['SchedShipDate'],
                                'TargetArrival': row['TargetArrival'],
                                'NotBefore': row['NotBefore'],
                                'ShipVia': row['ShipVia'],
                                'LastModifiedOn': row['LastModifiedOn']
                            }
                        )
                        skus.append(sku)
                    quantity = floor(quantity)  # convert to whole number for the rest of the SKUs

                for _ in range(int(abs(ceil(quantity)))):
                    sku = SKU(
                        id=invID,
                        bundleqty=row['Pcs/Bundle'],
                        width=row['Width_mm'],
                        height=row['Height_mm'],
                        length=newLength,
                        weight=row['Weight_kg'],
                        desc=row['Description'],
                        can_be_bottom=row['Can_be_bottom'],
                        data={
                            'OrderType': row['OrderType'],
                            'OrderNbr': row['OrderNbr'],
                            'UOM': row['UOM'],
                            'Bdl_Override': row['Bdl_Override'] if pd.notna(row['Bdl_Override']) else None,
                            'ShipTo': row['ShipTo'],
                            'AddressLine1': row['AddressLine1'],
                            'AddressLine2': row['AddressLine2'],
                            'City': row['City'],
                            'State': row['State'],
                            'Country': row['Country'],
                            'Status': row['Status'],
                            'OrderDate': row['OrderDate'],
                            'ProdReleaseDate': row['ProdReleaseDate'],
                            'SchedShipDate': row['SchedShipDate'],
                            'TargetArrival': row['TargetArrival'],
                            'NotBefore': row['NotBefore'],
                            'ShipVia': row['ShipVia'],
                            'LastModifiedOn': row['LastModifiedOn']
                        }
                    )
                    skus.append(sku)
            order_skus[order] = skus
        return order_skus

    def shrink_to_square(self, w, h, x, dim_to_shrink):
        """
        Shrinks the area of a rectangle by a multiplier `x`, changing only one dimension
        """
        if not (0 < x < 1):
            self.show_alert("Error", "Shrink multiplier must be between 0 and 1.", "error")

        original_area = w * h
        new_area = original_area * x

        if dim_to_shrink.lower() == 'height':
            # Option 2: change height, keep width
            new_h2 = new_area / w
            return (w, new_h2)
        elif dim_to_shrink.lower() == 'width':
            # Option 1: change width, keep height
            new_w1 = new_area / h
            return (new_w1, h)
        else:
            # shrink smaller dim if not specified
            if w < h:
                new_w1 = new_area / h
                return (new_w1, h)
            else:
                new_h2 = new_area / w
                return (w, new_h2)

    def write_optimized_bundles(self, workbook, order_bundles: dict):
        """
        Write the packed bundles to a new sheet in the workbook
        """
        # create a new sheet for the optimized bundles
        if "SO-PackExportData" in workbook.sheetnames:
            del workbook["SO-PackExportData"]
        if not self.append_data:
            if "Optimized_Bundles" in workbook.sheetnames:
                del workbook["Optimized_Bundles"]
            optimized_sheet = workbook.create_sheet("Optimized_Bundles")
        else:
            optimized_sheet = workbook["Optimized_Bundles"]

        # write headers
        intersect_headers = ['Can_be_bottom', 'Dim_shrink']
        # remove intersect headers from the main headers
        self.headers = [header for header in self.headers if header not in intersect_headers]
        # add headers
        self.headers.insert(6, 'TotalPcs')
        self.headers.insert(2, 'BundleNbr')
        self.headers.append('OptimizedOn')
        if not self.append_data:
            optimized_sheet.append(self.headers)

        top_groups = []
        bottom_groups = []

        # add data for each order's bundles
        for order, bundles in order_bundles.items():
            # add missing/removed skus as part of bundle "0"
            if order in [sku.data['OrderNbr'] for sku in self.missingDataSKUs]:
                # add a row for each missing SKU
                written_skus = set()  # to avoid writing the same SKU multiple times
                for sku in self.missingDataSKUs:
                    if sku.id in written_skus:
                        continue
                    else:
                        written_skus.add(sku.id)
                        # count the number of identical SKUs in the order
                        order_skus = [s for s in self.missingDataSKUs if s.id == sku.id]
                        quantity = len(order_skus)
                    if sku.data['OrderNbr'] == order:
                        optimized_sheet.append([
                            sku.data['OrderType'],
                            order,
                            0,
                            sku.data['Bdl_Override'],
                            sku.id,
                            quantity,
                            sku.bundleqty if sku.bundleqty else "N/A",  # default to 1 if bundleqty is None
                            "N/A" if not sku.bundleqty else quantity * sku.bundleqty,
                            "N/A",
                            "N/A",
                            "N/A",
                            "N/A",
                            sku.data['UOM'],
                            sku.desc,
                            sku.data['ShipTo'],
                            sku.data['AddressLine1'],
                            sku.data['AddressLine2'],
                            sku.data['City'],
                            sku.data['State'],
                            sku.data['Country'],
                            sku.data['Status'],
                            sku.data['OrderDate'].strftime("%Y-%m-%d") if sku.data['OrderDate'] else None,
                            sku.data['ProdReleaseDate'].strftime("%Y-%m-%d") if sku.data['ProdReleaseDate'] else None,
                            sku.data['SchedShipDate'].strftime("%Y-%m-%d") if sku.data['SchedShipDate'] else None,
                            sku.data['TargetArrival'],
                            sku.data['NotBefore'],
                            sku.data['ShipVia'],
                            sku.data['LastModifiedOn'].strftime("%Y-%m-%d") if sku.data['LastModifiedOn'] else None,
                            datetime.now().strftime("%Y-%m-%d"),
                        ])

            order_row_count = 0
            for bundle_index, bundle in enumerate(bundles):
                # get quantity of each SKU in the bundle (including stacked quantities)
                sku_counts = {}
                for sku in bundle.skus:
                    if sku.id not in sku_counts:
                        sku_counts[sku.id] = {'qty': 0, 'sku': sku}
                    sku_counts[sku.id]['qty'] += 1

                # calculate bundle actual dimensions and weight
                actual_width, actual_height, _ = bundle.get_actual_dimensions()
                total_weight = bundle.get_total_weight()

                packaging_skus_active = False
                packaging_idx = 2
                # write each SKU in the bundle to the sheet
                for sku_id, sku_data in sku_counts.items():
                    if "Pack_Angle" in sku_id and not packaging_skus_active:
                        packaging_skus_active = True
                        packaging_idx = optimized_sheet.max_row + 2
                    # check if data is None (this happens for Packaging SKUs)
                    if sku_data['sku'].data is None:
                        # give data from another SKU in the order, since they are the same
                        for _, nested_sku_data in sku_counts.items():
                            if nested_sku_data['sku'].data is not None:
                                sku_data['sku'].data = nested_sku_data['sku'].data
                                break
                    try:
                        optimized_sheet.append([
                            sku_data['sku'].data['OrderType'],
                            order,
                            bundle_index + 1,
                            sku_data['sku'].data['Bdl_Override'],
                            sku_id,
                            sku_data['qty'],
                            sku_data['sku'].bundleqty,
                            sku_data['qty'] * sku_data['sku'].bundleqty,
                            actual_width,
                            actual_height,
                            bundle.max_length,
                            total_weight,
                            sku_data['sku'].data['UOM'],
                            sku_data['sku'].desc,
                            sku_data['sku'].data['ShipTo'],
                            sku_data['sku'].data['AddressLine1'],
                            sku_data['sku'].data['AddressLine2'],
                            sku_data['sku'].data['City'],
                            sku_data['sku'].data['State'],
                            sku_data['sku'].data['Country'],
                            sku_data['sku'].data['Status'],
                            sku_data['sku'].data['OrderDate'].strftime("%Y-%m-%d") if sku_data['sku'].data['OrderDate'] else None,
                            sku_data['sku'].data['ProdReleaseDate'].strftime("%Y-%m-%d") if sku_data['sku'].data['ProdReleaseDate'] else None,
                            sku_data['sku'].data['SchedShipDate'].strftime("%Y-%m-%d") if sku_data['sku'].data['SchedShipDate'] else None,
                            sku_data['sku'].data['TargetArrival'],
                            sku_data['sku'].data['NotBefore'],
                            sku_data['sku'].data['ShipVia'],
                            sku_data['sku'].data['LastModifiedOn'].strftime("%Y-%m-%d") if sku_data['sku'].data['LastModifiedOn'] else None,
                            datetime.now().strftime("%Y-%m-%d"),
                        ])
                        order_row_count += 1
                    except Exception as e:
                        self.show_alert("Error", f"Error writing SKU {sku_id} to the sheet: {e}", "error")
                        return
                if packaging_skus_active:
                    bottom_groups.append([packaging_idx - 1, optimized_sheet.max_row])
            # group the rows by order number in Excel
            top_groups.append([optimized_sheet.max_row - order_row_count + 2, optimized_sheet.max_row])
            # add a blank row after each order's bundles
            optimized_sheet.append([])

        # create a table over the data
        if "OptimizedBundlesTable" in optimized_sheet.tables:
            del optimized_sheet.tables["OptimizedBundlesTable"]
        table = openpyxl.worksheet.table.Table(displayName="OptimizedBundlesTable", ref=optimized_sheet.dimensions, tableStyleInfo=openpyxl.worksheet.table.TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True))
        # resize the table to fit the data
        optimized_sheet.add_table(table)

        # if two blank rows in a row, remove the first one
        for row in range(optimized_sheet.max_row, 1, -1):
            # 2 consecutive blank rows
            if optimized_sheet[row][0].value is None and optimized_sheet[row - 1][0].value is None:
                optimized_sheet.delete_rows(row - 1, 1)
                optimized_sheet.row_dimensions[row + 1].outlineLevel = 1
            # if the row is blank, set the outline level to 0
            if optimized_sheet.row_dimensions[row].outlineLevel == 2:
                continue
            if optimized_sheet[row][0].value is None:
                optimized_sheet.row_dimensions[row].outlineLevel = 0
                optimized_sheet.row_dimensions[row + 1].outlineLevel = 0
            # if the row is not blank, set the outline level to 1
            else:
                optimized_sheet.row_dimensions[row].outlineLevel = 1
        optimized_sheet.row_dimensions[2].outlineLevel = 0  # first row is the header, no grouping

        for group in bottom_groups:
            # group packaging SKUs together
            optimized_sheet.row_dimensions.group(start=group[0], end=group[1], hidden=False)
            for row in range(group[0], group[1]):
                optimized_sheet.row_dimensions[row].outlineLevel = 2
        for row in range(1, optimized_sheet.max_row + 1):
            optimized_sheet.row_dimensions[row].hidden = False

        # save the workbook
        try:
            if self.append_data:
                workbook.save(self.ui.appendDir.text())
            else:
                workbook.save(f"{self.workingDir}/Optimized_Bundles.xlsx")
        except Exception as e:
            self.show_alert("Error", f"Error saving the file. Is it already open? Error: {e}", "error")
            return

    def show_alert(self, title, message, type="warning") -> None:
        """
        Show an alert dialog with the given title and message
        """
        if type == "warning":
            QMessageBox.warning(self.Widget, title, message)
        elif type == "info":
            QMessageBox.information(self.Widget, title, message)
        elif type == "error":
            QMessageBox.critical(self.Widget, title, message)
