import Model, View
import json, os, pyqtgraph as pg, numpy as np
from PyQt6 import QtWidgets
from PyQt6.QtCore import Qt
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter



class Controller():
    
    while True:
        try:
            pathFile = open("path.json")
            pathDict = json.load(pathFile)
            d = pathDict["etabs_path"]
            d = pathDict["model_path"]
            d = pathDict["model_name"]
            d = pathDict["folder_path"]
            pathFile.close()
        except (FileNotFoundError, json.decoder.JSONDecodeError,KeyError):
            a_dict = {"etabs_path": "Import path to etabs", "model_path": "Import path to model", "model_name": "Import model name",
                   "folder_path": "Import path to saving folder", "previously_connected_to": ""}
            #a_dict = {"etabs_path": "C:\\Program Files\\Computers and Structures\\ETABS 22",
             #       "model_path": "C:\\TestProject", "model_name": "test_model_rev2.8.EDB", "folder_path": "C:\\Report",
              #      "previously_connected_to": ""}
            with open('path.json', 'w') as outfile:
                json.dump(a_dict, outfile)
        else:
            break
    view = View.view()
    model =  Model.Model()
    def __init__(self):
            
        self.fillPathWindow()
        self.view.pathWindow.exitButton.pressed.connect(self.pathWindowExitButtonAction)
        self.view.pathWindow.closeButton.pressed.connect(self.pathWindowCloseButtonAction)
        self.view.pathWindow.runButton.pressed.connect(self.pathWindowRunButtonAction)
        self.view.pathWindow.ButtonEditEtabsPath.pressed.connect(self.ButtonEditEtabsPathAction)
        self.view.pathWindow.ButtonEditModelPath.pressed.connect(self.ButtonEditModelPathAction)
        self.view.pathWindow.ButtonEditSavePath.pressed.connect(self.ButtonEditSavePathAction)
        self.view.mainWindow.driftControlButton.pressed.connect(self.mainWindowDriftControlButtonAction)
        self.view.mainWindow.pathButton.pressed.connect(self.mainWindowpathButtonAction)
        self.view.mainWindow.closeButton.pressed.connect(self.mainWindowCloseButtenAction)
        self.view.mainWindow.connectButton.pressed.connect(self.mainWindowConnectButtonAction)
        self.view.driftControlWindow.closeButton.pressed.connect(self.driftControlWindowCloseACtions)
        self.view.driftControlWindow.listWidget.itemSelectionChanged.connect(self.sellectLoadCombAction)
        self.view.driftControlWindow.reortButton.pressed.connect(self.reportToExcel)
        self.view.driftControlWindow.my_signal.connect(self.mySlot)
        

        self.view.app.exec()
            

     
    
    def loadJson(self):
        pathFile = open("path.json")
        self.pathDict = json.load(pathFile)
        path = self.pathDict
        return path
        
        
    def saveJson(self):
        pathFile = open("path.json", "w")
        json.dump(self.pathDict, pathFile)
        return True
    
    def fillPathWindow(self):
        self.view.pathWindow.lineEditEtabsPath.setText(self.pathDict["etabs_path"])
        self.view.pathWindow.lineEditModelPath.setText(self.pathDict["model_path"])
        self.view.pathWindow.lineEditModelName.setText(self.pathDict["model_name"])
        self.view.pathWindow.lineEditSavePath.setText(self.pathDict["folder_path"])
        self.view.pathWindowForm.show()

            
    def pathWindowExitButtonAction(self):
        self.view.pathWindowForm.close()
        
    def pathWindowCloseButtonAction(self):
        self.pathDict["etabs_path"] = self.view.pathWindow.lineEditEtabsPath.text()
        self.pathDict["model_path"] = self.view.pathWindow.lineEditModelPath.text()
        self.pathDict["model_name"] = self.view.pathWindow.lineEditModelName.text()
        self.pathDict["folder_path"] = self.view.pathWindow.lineEditSavePath.text()
        self.saveJson()
        self.view.pathWindowForm.close()
            
    def pathWindowRunButtonAction(self):
        self.pathDict["etabs_path"] = self.view.pathWindow.lineEditEtabsPath.text()
        self.pathDict["model_path"] = self.view.pathWindow.lineEditModelPath.text()
        self.pathDict["model_name"] = self.view.pathWindow.lineEditModelName.text()
        self.pathDict["folder_path"] = self.view.pathWindow.lineEditSavePath.text()
        self.saveJson()
        self.view.pathWindowForm.close()
        if not os.path.exists(self.pathDict["etabs_path"] + os.sep + "ETABS.exe"):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("Etabs.exe was not founded at " + self.pathDict["etabs_path"])
            mesgBox.exec()
            self.fillPathWindow()
            return
        if not os.path.exists(self.pathDict["model_path"]):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("model path was not founded at " + self.pathDict["model_path"])
            mesgBox.exec()
            self.fillPathWindow()
            return
        if not os.path.exists(self.pathDict["model_path"]):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("model was not founded at " + self.pathDict["model_path"])
            mesgBox.exec()
            self.fillPathWindow()
            return
        if not os.path.exists(self.pathDict["folder_path"]):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("Report path was not founded at " + self.pathDict["folder_path"])
            mesgBox.exec()
            self.fillPathWindow()
            return
        self.saveJson()
        self.view.mainWindow.lineEditPreConnect.setText(str(self.pathDict["previously_connected_to"]))
        self.view.mainWindow.lineEditStatus.setText(str(self.pathDict["model_path"]))
        self.view.mainWindowForm.show()


    def driftControlWindowCloseACtions(self):
        self.view.driftControlWindow.emit_signal(self.update_label())
        self.view.mainWindowForm.show()
        self.view.driftControlWindowForm.close()


        
    def mainWindowDriftControlButtonAction(self):
        self.view.mainWindow.closeButton.setDisabled(True)
        self.view.mainWindow.connectButton.setDisabled(True)
        self.view.mainWindow.pathButton.setDisabled(True)
        self.view.mainWindow.driftControlButton.setDisabled(True)
        self.fillCombolist()
        self.view.mainWindow.closeButton.setEnabled(True)
        self.view.mainWindow.connectButton.setEnabled(True)
        self.view.mainWindow.pathButton.setEnabled(True)
        self.view.mainWindow.driftControlButton.setEnabled(True)
        self.view.driftControlWindow.line.setReadOnly = False
        self.view.driftControlWindowForm.show()
        self.view.mainWindowForm.hide()
        
    def mainWindowCloseButtenAction(self):
        self.view.mainWindowForm.close()


    def mainWindowpathButtonAction(self):
        self.view.mainWindow.driftControlButton.setDisabled(True)
        self.view.mainWindow.lineEditPreDrift.setText("")
        self.view.driftControlWindow.listWidget.clear()
        self.view.driftControlWindow.tableWidget.clear()
        self.view.driftControlWindow.listWidget.scrollToTop()
        self.view.driftControlWindow.listWidget.clear()
        self.view.driftControlWindow.line.clear()
        self.view.driftControlWindow.plot.clear()
        self.view.mainWindowForm.close()
        self.fillPathWindow()
        
    def mainWindowConnectButtonAction(self):
        self.view.mainWindow.closeButton.setDisabled(True)
        self.view.mainWindow.connectButton.setDisabled(True)
        self.view.mainWindow.pathButton.setDisabled(True)
        self.view.mainWindow.driftControlButton.setDisabled(True)
        mesgBox = QtWidgets.QMessageBox()
        mesgBox.setText("please don't close etabs")
        mesgBox.exec()
        self.model.runModel()
        self.view.mainWindow.closeButton.setEnabled(True)
        self.view.mainWindow.connectButton.setEnabled(True)
        self.view.mainWindow.pathButton.setEnabled(True)
        self.view.mainWindow.driftControlButton.setEnabled(True)
        self.pathDict["previously_connected_to"] = self.pathDict["model_path"]
        pathFile = open("path.json", "w")
        json.dump(self.pathDict, pathFile)

        
    def selectDirectory(self):
        self.dialog = QtWidgets.QFileDialog()
        self.folder_path = self.dialog.getExistingDirectory(None, "Select Folder")
        return self.folder_path


    def selectFile(self):
        self.dialog = QtWidgets.QFileDialog(caption="Choose File")
        file_filter = "ETABS Models (*.EDB)"
        self.filedir = self.dialog.getOpenFileNames(None, "Choose File", "", file_filter)
        self.name = self.filedir[0][0].split("/")[-1]
        return [self.name, self.filedir[0][0]]
        

    def ButtonEditEtabsPathAction(self):
        self.path = self.selectDirectory()
        if not self.path == "":
            self.view.pathWindow.lineEditEtabsPath.setText(self.path)

    def ButtonEditModelPathAction(self):
        path = self.selectFile()
        if not path[0] == "":
            self.view.pathWindow.lineEditModelName.setText(path[0])
            self.view.pathWindow.lineEditModelPath.setText(path[1])
        
    def ButtonEditSavePathAction(self):
        path = self.selectDirectory()
        self.view.pathWindow.lineEditSavePath.setText(path)
        
    def fillCombolist(self):        
        combosName = self.model.ComboName()
        index = 0
        for comboName in combosName:
            item = QtWidgets.QListWidgetItem()
            self.view.driftControlWindow.listWidget.addItem(item)
            item = self.view.driftControlWindow.listWidget.item(index)
            index += 1
            item.setText(str(comboName))
            
            
    def fillComboTableAndLine(self):
        
        self.view.driftControlWindow.tableWidget.setRowCount(len(self.storyslist))
        self.view.driftControlWindow.tableWidget.setColumnCount(5)
        self.view.driftControlWindow.tableWidget.setColumnWidth(0, 25)
        self.view.driftControlWindow.tableWidget.setColumnWidth(1, 65)
        self.view.driftControlWindow.tableWidget.setColumnWidth(2, 65)
        self.view.driftControlWindow.tableWidget.setColumnWidth(3, 65)
        self.view.driftControlWindow.tableWidget.setColumnWidth(4, 65)
        self.view.driftControlWindow.tableWidget.setHorizontalHeaderLabels(["Sty", "1000XDisp", "1000YDisp","1000XDrift" ,"1000YDrift"])
        self.maxDispX = 0
        self.maxDispY = 0
        
        counter = 0
        for story in self.storyslist:
            try:
                XDrift = self.dfTable[( self.dfTable["Story"] == story)]['1000DriftX'].max()
            except(IndexError):
                XDrift = 0.0    
                
            try:
                YDrift =  self.dfTable[( self.dfTable["Story"] == story)]['1000DriftY'].max()
            except(IndexError):
                YDrift = 0.0
                
            try:
                Xdisp =  self.dfTable[( self.dfTable["Story"] == story)]['1000DispX'].mean()
            except(IndexError):
                Xdisp = 0.0  
                
            try:
                Ydisp =  self.dfTable[( self.dfTable["Story"] == story)]['1000DispY'].mean()
            except(IndexError):
                Ydisp = 0.0  
                
            if abs(self.maxDispX) < abs(Xdisp):
                self.maxDispX = Xdisp
                
            if abs(self.maxDispY) < abs(Ydisp):
                self.maxDispY = Ydisp
                
            item_story = QtWidgets.QTableWidgetItem(str(story))
            item_story.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_XDisp = QtWidgets.QTableWidgetItem(str("%.5f" % Xdisp))
            item_XDisp.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_YDisp = QtWidgets.QTableWidgetItem(str("%.5f" % Ydisp))
            item_YDisp.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_XDrift = QtWidgets.QTableWidgetItem(str("% 5f" % XDrift))
            item_XDrift.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_YDrift = QtWidgets.QTableWidgetItem(str("%.5f" % YDrift))
            item_YDrift.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.view.driftControlWindow.tableWidget.setItem(counter, 0, item_story)
            self.view.driftControlWindow.tableWidget.setItem(counter, 1, item_XDisp)
            self.view.driftControlWindow.tableWidget.setItem(counter, 2, item_YDisp)
            self.view.driftControlWindow.tableWidget.setItem(counter, 3, item_XDrift)
            self.view.driftControlWindow.tableWidget.setItem(counter, 4, item_YDrift)
            counter += 1
            


    def fillGraphForCombo(self):
        self.view.driftControlWindow.plot.clear()
        ListX = []
        ListY = []
        for story in self.storyslist:
            try:
                XDrift = self.dfTable[( self.dfTable["Story"] == story)]['1000DriftX'].max()
            except(IndexError):
                XDrift = 0.0    
                
            try:
                YDrift =  self.dfTable[( self.dfTable["Story"] == story)]['1000DriftY'].max()
            except(IndexError):
                YDrift = 0.0
            ListX.append(XDrift)
            ListY.append(YDrift)
        
        self.storyelev = self.model.storyElev()
        self.storyelev.reverse()
        
        XDriftCritical = []
        XHCritical = []
        YDriftCritical = []
        YHCritical = []
        for i in range(len(ListX)):
            if ListX[i] > self.calcDriftThreshold():
                XDriftCritical.append(ListX[i])
                XHCritical.append(self.storyelev[i])
            
        for i in range(len(ListY)):
            if ListY[i] > self.calcDriftThreshold():
                YDriftCritical.append(ListY[i])
                YHCritical.append(self.storyelev[i])
                
        
        npXDriftCritical = np.array(XDriftCritical)
        npXHCritical = np.array(XHCritical)
        npYDriftCritical = np.array(YDriftCritical)
        npYHCritical = np.array(YHCritical)        


        # Create a ScatterPlotItem and add it to the PlotWidget
        self.view.driftControlWindow.plot.showGrid(x=True, y=True, alpha=0.5)
        self.view.driftControlWindow.plot.plot(self.storyelev,ListY, pen = "b", symbolPen ='b', symbol ='o', symbolSize = 8, name ="Drift Y")
        self.view.driftControlWindow.plot.plot(self.storyelev,ListX, pen = "r", symbolPen ='r', symbol ='x', symbolSize = 8, name ="Drift X")
        scatterX = pg.ScatterPlotItem(x=npXHCritical, y=npXDriftCritical, size=10, brush="m", name = "X-Dir Critical Drift")
        self.view.driftControlWindow.plot.addItem(scatterX)
        
        scatterY = pg.ScatterPlotItem(x=npYHCritical, y=npYDriftCritical, size=10, brush="c",  name = "X-Dir Critical Drift")
        self.view.driftControlWindow.plot.addItem(scatterY)
        
    
    def sellectLoadCombAction(self):
        try:
            self.dfTable = self.model.storyDispForCombo(self.view.driftControlWindow.listWidget.currentItem().text()) 
            self.dfTable.sort_index(inplace=True)
            self.storyslist = list(dict.fromkeys(self.dfTable["Story"].to_list()))    
            self.fillGraphForCombo()  
            self.view.driftControlWindow.label_2.setText("Critcal Drift is " + str(self.calcDriftThreshold()) + "mm (Not true only fo cheking the app)")
            self.fillComboTableAndLine()               
        except Exception as e:
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("DisConneted from etabs Connect again.")
            mesgBox.exec()
        
        
    def reportToExcel(self):

        if self.view.driftControlWindow.line.text() == "":
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("Reort Name is Empty")
            mesgBox.exec()
            return
        
        try:        
            drift_threshold = self.calcDriftThreshold()
            wb = Workbook()
            ws = wb.active
            ws.title = "Drift Analysis Report"
            ws.cell(row=1, column=1, value="Story")
            ws.cell(row=1, column=2, value='DispX (mm)')
            ws.cell(row=1, column=3, value="DispY (mm)")
            ws.cell(row=1, column=4, value="DriftX (mm)")
            ws.cell(row=1, column=5, value="DriftY (mm)")

            rowIndex = 2
            for story in self.storyslist:
                try:
                    XDrift = self.dfTable[( self.dfTable["Story"] == story)]['1000DriftX'].max()
                except(IndexError):
                    XDrift = 0.0    
                    
                try:
                    YDrift =  self.dfTable[( self.dfTable["Story"] == story)]['1000DriftY'].max()
                except(IndexError):
                    YDrift = 0.0
                    
                try:
                    Xdisp =  self.dfTable[( self.dfTable["Story"] == story)]['1000DispX'].mean()
                except(IndexError):
                    Xdisp = 0.0  
                    
                try:
                    Ydisp =  self.dfTable[( self.dfTable["Story"] == story)]['1000DispY'].mean()
                except(IndexError):
                    Ydisp = 0.0  
                    
                
                ws.cell(row=rowIndex, column=1, value=story)
                ws.cell(row=rowIndex, column=2, value=Xdisp)
                ws.cell(row=rowIndex, column=3, value=Ydisp)
                ws.cell(row=rowIndex, column=4, value=XDrift)
                ws.cell(row=rowIndex, column=5, value=YDrift)
                
                rowIndex += 1
                
            # Apply Conditional Formatting for Drift X and Drift Y Columns
            drift_x_col = 4  # Column for Drift X
            drift_y_col = 5  # Column for Drift Y

            # Add a ColorScaleRule for highlighting drift values
            color_scale = ColorScaleRule(
                start_type="num", start_value=drift_threshold*0.7, start_color="FFFFFFFF",  # Green for low values
                mid_type="num", mid_value=drift_threshold *0.85, mid_color="FFFFFF00",  # Yellow
                end_type="num", end_value=drift_threshold, end_color="FFFF0000"  # Red for high values
            )

            # Apply the rule to Drift X and Drift Y
            ws.conditional_formatting.add(f"{get_column_letter(drift_x_col)}2:{get_column_letter(drift_x_col)}{len(self.storyslist)+1}", color_scale)
            ws.conditional_formatting.add(f"{get_column_letter(drift_y_col)}2:{get_column_letter(drift_y_col)}{len(self.storyslist)+1}", color_scale)

        
        except(AttributeError, ValueError, TypeError):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("No selected Load or discanected From Etabs")
            mesgBox.exec()
            return

        try:
            reportFileName = self.view.driftControlWindow.line.text()
            wb.save(self.pathDict["folder_path"] + "/" +reportFileName + ".xlsx")
        except(PermissionError):
            mesgBox = QtWidgets.QMessageBox()
            mesgBox.setText("Ther is an open file in " + self.pathDict["folder_path"] + "\\" + reportFileName + " close the file or change the report name")
            mesgBox.exec()
            
        mesgBox = QtWidgets.QMessageBox()
        mesgBox.setText("The reort file is saved in " + self.pathDict["folder_path"] + "\\" + reportFileName + ".")
        mesgBox.exec()
        
        
    def calcDriftThreshold(self):
        
        if len(self.storyelev) < 5:
            return 1000*0.025*max(self.storyelev)/400
        
        else:
            return 1000*0.02*max(self.storyelev)/400
        

        
    def update_label(self):
        # Update the label with the received text
        preText = "for load " + self.view.driftControlWindow.listWidget.currentItem().text() + "Max Displacement for X=" + str("%.3f" % self.maxDispX) + "mm and for Y=" + str("%.3f" % self.maxDispY) + "mm"
        return preText
    
    def mySlot(self, text):
        self.view.mainWindow.lineEditPreDrift.setText(text)