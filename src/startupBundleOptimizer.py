import sys
from PyQt6 import QtWidgets
import BundleGUI as gui

def handleException():
    """
    Handle exceptions by displaying the error message and exiting the application.
    """
    def excepthook(type, value, traceback):
        gui.excepthook(type, value, traceback)
    return excepthook

def startGUI():
    # start the Qt GUI application
    app = QtWidgets.QApplication(sys.argv)

    # Fusion style
    app.setStyle("Fusion")

    # create an instance of the GUI
    ui = gui.ProgramGUI()

    # exit the GUI
    sys.exit(app.exec())

if __name__ == "__main__":
    # set the exception hook to handle uncaught exceptions
    sys.excepthook = handleException()

    # start the GUI application
    startGUI()