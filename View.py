import sys
import DriftControlWindow, MainWindow, PathWindow
from PyQt6 import QtWidgets

class view():
    def __init__(self):
        self.app = QtWidgets.QApplication(sys.argv)    
        self.pathWindow = PathWindow.PathWindowUi_Form()
        self.pathWindowForm = QtWidgets.QWidget()
        self.pathWindow.setupUi(self.pathWindowForm)

        self.mainWindow = MainWindow.MainWindowUi_Form()
        self.mainWindowForm = QtWidgets.QWidget()
        self.mainWindow.setupUi(self.mainWindowForm)

        self.driftControlWindow = DriftControlWindow.DriftControlWindowUi_Form()
        self.driftControlWindowForm = QtWidgets.QWidget()
        self.driftControlWindow.setupUi(self.driftControlWindowForm)

