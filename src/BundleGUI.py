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

# Define filler materials
FILLER_44 = SKU(
    id="Pack_44Filler",
    width=100,
    height=100,
    length=3660,
    weight=1.810,
    desc="Pack 44 Filler Material"
)

FILLER_62 = SKU(
    id="Pack_62Filler", 
    width=150,
    height=50,
    length=3660,
    weight=2.268,
    desc="Pack 62 Filler Material"
)

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


    def get_workbook(self):
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
        self.data = self.get_data(self.workbook)
        self.ui.excelDir.setText(path)

    def get_data(self, workbook):
        """
        Read data from the 'SO_Input' sheet of the workbook
        """
        # get the "SO_Input" sheet
        if "SO_Input" not in workbook.sheetnames:
            self.showAlert("Warning", "Sheet 'SO_Input' not found in the selected file.")
            return None
        sheet = workbook["SO_Input"]

        # create headers with dataframe
        headers = [cell.value for cell in sheet[1]] # first row as headers
        df = pd.DataFrame(columns=headers)
        # read all rows from the sheet
        for row in sheet.iter_rows(min_row=2, values_only=True):
            df.loc[len(df)] = row  # append each row to the dataframe

        return df

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
        if self.data.empty:
            self.showAlert("Error", "Please select an Excel file first.", "error")
            return
        data = self.data
        self.ui.progressBar.setValue(10)
        self.ui.progressLabel.setText("Getting data...")

        data = self.updateQuantities(data)
        if data.empty:
            return

        # get unique orders
        unique_orders = data['Order ID'].unique()

        # get array of rows in sorted_data that belong to each order
        order_rows = {order: data[data['Order ID'] == order] for order in unique_orders}

        # create SKU objects for each order
        order_skus = self.create_sku_objects(order_rows)

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

        # write the packed bundles to a new sheet in the workbook
        self.write_optimized_bundles(self.workbook, order_bundles)

    def updateQuantities(self, data: pd.DataFrame) -> pd.DataFrame:
        """
        Open the "SKU_Quantities" sheet and get the sub-bundle quantities
        """
        if "SKU_Quantities" not in self.workbook.sheetnames:
            self.showAlert("Warning", "Sheet 'SKU_Quantities' not found in the selected file.")
            return {}

        sku_quantities_sheet = self.workbook["SKU_Quantities"]
        sku_quantities = pd.DataFrame(sku_quantities_sheet.values, columns=[cell.value for cell in sku_quantities_sheet[1]])

        # the quantities in 'data' must be MULTIPLES of the quantities in 'SKU_Quantities'
        for _, row in sku_quantities.iterrows():
            sku_id = row['SKU']
            if sku_id in data['SKU'].values:
                # get the quantity of the SKU in the data, order IDs must match
                if 'Order ID' in data.columns:
                    order_id = row['Order ID']
                    sku_qty = data.loc[(data['SKU'] == sku_id) & (data['Order ID'] == order_id), 'SKU Qty'].values[0]
                else:
                    continue
                if order_id == "SO-1013178" and sku_id == "6VR.289.15DWL":
                    pass
                # get the quantity of the SKU in the SKU_Quantities sheet
                sub_bundle_qty = row['Total Quantity']
                # calculate the new quantity as a multiple of the sub-bundle quantity
                new_qty = ceil(sku_qty / sub_bundle_qty)
                data.loc[(data['SKU'] == sku_id) & (data['Order ID'] == order_id), 'SKU Qty'] = new_qty * sub_bundle_qty

        return data

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
                        width=row['SKU Width (mm)'],
                        height=row['SKU Height (mm)'],
                        length=row['SKU Length (mm)'],
                        weight=row['SKU Weight (kg)'],
                        desc=row['SKU Description']
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
                   "Qty",
                   "SKU Description",
                   "Bundle Actual Width (mm)",
                   "Bundle Actual Height (mm)",
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
                        sku_counts[sku.id] = {'qty': 0, 'description': sku.desc, 'weight': sku.weight}
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
                        sku_data['description'],
                        actual_width,
                        actual_height,
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
