import sys
from PyQt6 import QtWidgets
from BundleGUI import Ui_MainWindow

# start the Qt GUI application
app = QtWidgets.QApplication(sys.argv)

# Fusion style
app.setStyle("Fusion")

# create an instance of the GUI
ui = Ui_MainWindow()

# exit the GUI
sys.exit(app.exec())
