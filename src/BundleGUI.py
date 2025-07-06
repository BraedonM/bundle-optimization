from PyQt6.QtWidgets import QMainWindow, QApplication, QMessageBox, QWidget, QFileDialog
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap
from BundleQtGui import Ui_BundleOptimizer

import openpyxl
import pandas as pd
from tkinter import filedialog
from typing import List
import os
from itertools import groupby
from math import ceil

from bundle_classes import SKU, Bundle
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
        # Initialize values and connect signals
        self.ui.fileBrowse.clicked.connect(self.get_workbook)
        self.ui.openExample.clicked.connect(self.openExampleFile)
        self.ui.optimizeBundles.clicked.connect(self.optimizeBundles)
        self.ui.openImages.clicked.connect(self.openImages)
        self.ui.openExcel.clicked.connect(self.openExcel)
        self.ui.helpButton.clicked.connect(self.openHelp)

        self.data = pd.DataFrame()
        self.maxWidth = 590
        self.maxHeight = 590
        self.maxLength = 3680
        self.workingDir = None  # to hold the directory of the selected Excel file
        self.missingDataSKUs = []  # to hold SKUs that are missing data in the Excel file


    def get_input_workbook(self):
        """
        Open a dialog to pick an Excel file, return the workbook object
        """
        # get excel file path from user
        path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx; *.xls; *.xlsm")],
        )
        if not path:
            return None

        self.workingDir = path.rsplit('/', 1)[0]  # get the directory of the selected file

        # load the workbook and read the sheet "SO_Input"
        self.workbook = openpyxl.load_workbook(path)
        self.ui.excelDir.setText(path)

    def get_sub_bundle_data_sheet(self):
        """
        Open a dialog to pick an Excel file, return the sheet "Sub-Bundle_Data"
        """
        # load the workbook and read the sheet "Sub-Bundle_Data"
        path = os.path.join(os.path.dirname(__file__), 'Sub-Bundle_Data.xlsx')
        workbook = openpyxl.load_workbook(path)
        sheet = workbook["Sub-Bundle_Data"]
        if not sheet:
            self.showAlert("Warning", "Sheet 'Sub-Bundle_Data' not found in file.")
            return None
        return sheet

    def get_data(self, workbook):
        """
        Read data from the 'SO_Input' sheet of the workbook
        """
        # get the "SO_Input" sheet
        if "SO_Input" not in workbook.sheetnames:
            self.showAlert("Warning", "Sheet 'SO_Input' not found in the selected file.")
            return None
        so_input = workbook["SO-PackExportData"]
        if not so_input:
            self.showAlert("Warning", "Sheet 'SO-PackExportData' is empty or not found. Using first sheet in the file instead.")
            so_input = workbook.active  # fallback to the first sheet if 'SO_Input' is not found
        sb_data = self.get_sub_bundle_data_sheet()

        df = pd.DataFrame(columns=[cell.value for cell in so_input[1]])
        sb_df = pd.DataFrame(columns=[cell.value for cell in sb_data[2]])
        # read all rows from the sheets
        for row in so_input.iter_rows(min_row=2, values_only=True):
            df.loc[len(df)] = row
        for row in sb_data.iter_rows(min_row=3, values_only=True):
            # check for any empty cells in the row
            if any(cell is None for cell in row):
                # if row[0] not in self.missingDataSKUs:
                #     self.missingDataSKUs.append(row[0])  # assuming SKU is in the second column
                continue
            sb_df.loc[len(sb_df)] = row

        # remove any rows with SKUs that are in the missingDataSKUs list
        df = df[~df['SKU'].isin(self.missingDataSKUs)]

        # check if the required columns are present
        df_cols = "Order ID", "SKU", "SKU Qty", "SKU Qty/Bundle", "SKU Description", "SKU Width (mm)", "SKU Height (mm)", "SKU Length (mm)", "SKU Weight (kg)"
        for col in df_cols:
            if col not in df.columns:
                # add the column with default values
                df[col] = None

        # put data from sb_df into df rows
        for _, row in sb_df.iterrows():
            sku_id = row['SKU']
            df_rows = []
            for index, df_row in df.iterrows():
                if sku_id in df_row['SKU']:
                    df_rows.append(index)
            for df_row_idx in df_rows:
                # update the SKU Qty/Bundle column with the value from sb_df
                df.loc[df_row_idx, 'SKU Qty/Bundle'] = row['Qty/bundle']
                df.loc[df_row_idx, 'SKU Width (mm)'] = row['Width (mm)']
                df.loc[df_row_idx, 'SKU Height (mm)'] = row['Height (mm)']
                df.loc[df_row_idx, 'SKU Length (mm)'] = row['Length (mm)']
                df.loc[df_row_idx, 'SKU Weight (kg)'] = row['Weight kg/length']

        # iterate through df and convert qty (in pieces) to qty (in bundles)
        for index, row in df.iterrows():
            if row['SKU Qty/Bundle'] is not None and row['SKU Qty'] is not None:
                try:
                    # convert SKU Qty to SKU Qty/Bundle
                    df.loc[index, 'SKU Qty'] = ceil(row['SKU Qty'] / row['SKU Qty/Bundle'])
                except ZeroDivisionError:
                    self.showAlert("Error", f"SKU Qty/Bundle cannot be zero for SKU {row['SKU']}. Please check the input data.", "error")
                    return pd.DataFrame()

        return df

    def removeInvalids(self, order_skus: dict) -> dict:
        """
        Iterate through the order_skus dictionary and remove any SKUs that have None as width, height, length, or weight.
        """
        for order, skus in order_skus.items():
            valid_skus = []
            for sku in skus:
                if sku.width is None or sku.height is None or sku.length is None or sku.weight is None:
                    if sku.id not in self.missingDataSKUs:
                        self.missingDataSKUs.append(sku.id)  # add to missing data SKUs list
                else:
                    valid_skus.append(sku)
            order_skus[order] = valid_skus
        return order_skus

    def openExampleFile(self):
        """
        Open the example Excel file to the user, in Excel
        """
        example_path = os.path.join(os.path.dirname(__file__), 'SO_Input_Example.xlsx')
        if os.path.exists(example_path):
            os.startfile(example_path)
        else:
            self.showAlert("File Not Found", "Example file not found.")

    def optimizeBundles(self):
        """
        Optimize bundles based on the data from the Excel file
        """
        self.ui.progressLabel.setText("Getting data...")
        self.ui.progressBar.setValue(10)
        data = self.get_data(self.workbook)
        if data.empty:
            self.showAlert("Error", "Invalid path for Excel file.", "error")
            return

        # get unique orders
        unique_orders = data['Order ID'].unique()

        # get array of rows in sorted_data that belong to each order
        order_rows = {order: data[data['Order ID'] == order] for order in unique_orders}

        # create SKU objects for each order
        order_skus = self.create_sku_objects(order_rows)

        order_skus = self.removeInvalids(order_skus)

        self.ui.progressBar.setValue(20)
        images_dir = f"{self.workingDir}/images"
        os.makedirs(images_dir, exist_ok=True)

        images_dir = f"{self.workingDir}/images"
        os.makedirs(images_dir, exist_ok=True)

        # pack each order's SKUs into bundles
        order_bundles = {}
        for order, skus in order_skus.items():
            bundles = pack_skus(skus, self.maxWidth, self.maxHeight)
            order_bundles[order] = bundles

            visualize_bundles(bundles, f"{images_dir}/Order_{order}.png")
            self.ui.progressBar.setValue(int(round(20 + 80 * (list(order_skus.keys()).index(order) + 1) / len(order_skus))))
            self.ui.progressLabel.setText(f"Packing order {order}...")

        self.images_dir = images_dir
        self.ui.progressBar.setValue(100)
        self.ui.progressLabel.setText("Packing complete!")

        if self.missingDataSKUs:
            self.showAlert("Missing Data", f"The following SKUs are missing data and could not\nbe included in the optimization:\n\n{'\n'.join(self.missingDataSKUs)}\n\nPlease check the input data.", "warning")

        # write the packed bundles to a new sheet in the workbook
        self.write_optimized_bundles(self.workbook, order_bundles)

    def create_sku_objects(self, order_rows: dict):
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
                        bundleqty=row['SKU Qty/Bundle'],
                        width=row['SKU Width (mm)'],
                        height=row['SKU Height (mm)'],
                        length=row['SKU Length (mm)'],
                        weight=row['SKU Weight (kg)'],
                        desc=row['SKU Description'],
                        can_be_bottom=['Bottom Row Acceptable']
                    )
                    skus.append(sku)
            order_skus[order] = skus
        return order_skus


    def write_optimized_bundles(self, workbook, order_bundles: dict):
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
                   "Sub-Bundle Qty",
                   "Pcs/Sub-Bundle",
                   "Total Pcs",
                   "SKU Description",
                   "Bundle Width (mm)",
                   "Bundle Height (mm)",
                   "Bundle Length (mm)",
                   "Bundle Total Weight (kg)",
                   "Note"
                   ]
        optimized_sheet.append(headers)

        # add data for each order's bundles
        for order, bundles in order_bundles.items():
            for bundle_index, bundle in enumerate(bundles):
                # get quantity of each SKU in the bundle (including stacked quantities)
                sku_counts = {}
                for sku in bundle.skus:
                    if sku.id not in sku_counts:
                        sku_counts[sku.id] = {'qty': 0, 'description': sku.desc, 'weight': sku.weight, 'bundleqty': sku.bundleqty}
                    sku_counts[sku.id]['qty'] += sku.stacked_quantity
                
                # calculate bundle actual dimensions and weight
                actual_width, actual_height, _ = bundle.get_actual_dimensions()
                total_weight = bundle.get_total_weight()
                
                # write each SKU in the bundle to the sheet
                for sku_id, sku_data in sku_counts.items():
                    optimized_sheet.append([
                        order,
                        bundle_index + 1,
                        sku_id,
                        sku_data['qty'],
                        sku_data['bundleqty'],
                        sku_data['qty'] * sku_data['bundleqty'],
                        sku_data['description'],
                        actual_width,
                        actual_height,
                        bundle.max_length,
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
            workbook.save(f"{self.workingDir}/Optimized_Bundles.xlsx")
        except Exception as e:
            self.showAlert("Error", f"Error saving the file. Is it already open? Error: {e}", "error")
            return

    def openImages(self):
        """
        Open the images directory in the file explorer
        """
        if hasattr(self, 'images_dir'):
            os.startfile(self.images_dir)
        else:
            self.showAlert("No Images", "No images have been generated yet. Please optimize bundles first.", "error")

    def openExcel(self):
        """
        Open the optimized bundles Excel file in the file explorer
        """
        if not self.workingDir:
            self.showAlert("Error", "No file found", "error")
            return
        excel_path = f"{self.workingDir}/Optimized_Bundles.xlsx"
        if os.path.exists(excel_path):
            os.startfile(excel_path)
        else:
            self.showAlert("Error", "Optimized bundles file not found. Please optimize bundles first.", "error")

    def openHelp(self):
        """
        Open the help file in the file explorer
        """
        self.showAlert("Help", """
        For additional support, please contact Aionex Solutions Ltd.
        (604-309-8975)

        To use this tool, follow these steps:

        1. Click on 'Browse' to select your Excel file containing SKU data.
            - If you are unsure how to format your Excel file,
              you can click on 'Open Example File' to open a sample file.
            - Please format your Excel file according to the example provided.

        2. Click on 'Perform Bundle Optimization' to start the
           optimization process.
           Wait for the progress bar to complete.

        3. Once the optimization is complete, you can:
            - Click on 'Open Images Folder' to view the generated bundle images.
            - Click on 'Open Resultant Excel File' to view the optimized
              bundle data in an Excel file.

        NOTE: All files are generated in the same directory as the input Excel file.
        """, type="info")

    def showAlert(self, title, message, type="warning") -> None:
        """
        Show an alert dialog with the given title and message
        """
        if type == "warning":
            QMessageBox.warning(self.Widget, title, message)
        elif type == "info":
            QMessageBox.information(self.Widget, title, message)
        elif type == "error":
            QMessageBox.critical(self.Widget, title, message)
