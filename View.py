import json, sys
import DriftControlWindow, MainWindow, PathWindow
from PyQt6 import QtWidgets

pathFile = open("path.json")
pathDict = json.load(pathFile)

class View():

    
    pathFile = open("path.json")
    pathDict = json.load(pathFile)
    pathFile.close()
    app = QtWidgets.QApplication(sys.argv)    
    lasttext = ""

    pathWindow = PathWindow.Ui_Form()
    pathWindowForm = QtWidgets.QWidget()
    pathWindow.setupUi(pathWindowForm)

    mainWindow = MainWindow.Ui_Form()
    mainWindowForm = QtWidgets.QWidget()
    mainWindow.setupUi(mainWindowForm)


    driftControlWindow = DriftControlWindow.Ui_Form()
    driftControlWindowForm = QtWidgets.QWidget()
    driftControlWindow.setupUi(driftControlWindowForm)

    def fillPathWindow(self):
        self.pathWindow.lineEditEtabsPath.setText(pathDict["etabs_path"])
        self.pathWindow.lineEditModelPath.setText(pathDict["model_path"])
        self.pathWindow.lineEditModelName.setText(pathDict["model_name"])
        self.pathWindow.lineEditSavePath.setText(pathDict["folder_path"])
        self.pathWindowForm.show()

        
    def pathWindowExitButtonAction(self):
        self.pathWindowForm.close()
        
    def pathWindowCloseButtonAction(self):
        pathDict["etabs_path"] = self.pathWindow.lineEditEtabsPath.text()
        pathDict["model_path"] = self.pathWindow.lineEditModelPath.text()
        pathDict["model_name"] = self.pathWindow.lineEditModelName.text()
        pathDict["folder_path"] = self.pathWindow.lineEditSavePath.text()
        pathFile = open("path.json", "w")
        json.dump(pathDict, pathFile)
        self.pathWindowForm.close()
        
    def pathWindowRunButtonAction(self):
        pathDict["etabs_path"] = self.pathWindow.lineEditEtabsPath.text()
        pathDict["model_path"] = self.pathWindow.lineEditModelPath.text()
        pathDict["model_name"] = self.pathWindow.lineEditModelName.text()
        pathDict["folder_path"] = self.pathWindow.lineEditSavePath.text()
        self.mainWindow.lineEditPreConnect.setText(str(" "+pathDict["previously_connected_to"]))
        self.mainWindow.lineEditStatus.setText(str(pathDict["model_path"] +"\\"+ pathDict["model_name"]))
        pathFile = open("path.json", "w")
        json.dump(pathDict, pathFile)
        self.mainWindowForm.show()
        self.pathWindowForm.close()

    def driftControlWindowCloseACtions(self):
        #_str = "pre " + str(self.driftControlWindow.line.text())
        self.mainWindow.lineEditPreDrift.setText(_str)
        self.driftControlWindow.line.setText("")
        self.driftControlWindowForm.close()
        self.mainWindowForm.show()
        
    def mainWindowDriftControlButtonAction(self):
        #_str = "time is= " + str(datetime.datetime.now())
        self.driftControlWindow.line.setText(_str)
        self.driftControlWindowForm.show()
        self.mainWindowForm.hide()
        
    def mainWindowCloseButtenAction(self):
        self.mainWindowForm.close()


    def mainWindowpathButtonAction(self):
        self.mainWindowForm.hide()
        self.fillPathWindow()
        
    def mainWindowConnectButtonAction(self):
        text = (str(pathDict["model_path"] +"\\"+ pathDict["model_name"]))
        pathDict["previously_connected_to"] = text
        pathFile = open("path.json", "w")
        json.dump(pathDict, pathFile)
        
