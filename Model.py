import os, comtypes.client, json
import pandas as pd


class Model():

    
    def getModel(self):
        try:
            self.helper = comtypes.client.CreateObject('ETABSv1.Helper')
            self.helper = self.helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
            self.myETABSObject = self.helper.GetObject("CSI.ETABS.API.ETABSObject")
            self.SapModel = self.myETABSObject.SapModel
            self.ret = self.SapModel.Analyze.RunAnalysis()               
            return True     
        except (TypeError, AttributeError) as e:
            return False
        
        


    def runModel(self):
        with open("path.json", "r") as self.pathFile:
            self.pathDict = json.load(self.pathFile)
        self.helper = comtypes.client.CreateObject('ETABSv1.Helper')
        self.helper = self.helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

        try:
            #get the active ETABS object
            self.myETABSObject = self.helper.GetObject("CSI.ETABS.API.ETABSObject")
            self.SapModel = self.myETABSObject.SapModel
            self.ret = self.SapModel.Analyze.RunAnalysis()
        except (OSError, comtypes.COMError, AttributeError):
            self.myETABSObject = self.helper.CreateObject(self.pathDict["etabs_path"] + os.sep + "ETABS.exe")
            self.myETABSObject.ApplicationStart()
            self.SapModel = self.myETABSObject.SapModel
            self.ModelPath = self.pathDict["model_path"]
            self.SapModel.InitializeNewModel()
            self.ret = self.SapModel.File.OpenFile(self.ModelPath)
            self.ret = self.SapModel.Analyze.RunAnalysis()
        
    ConnectionsTATUS = True
    def ExitEtabs(self):
        self.myETABSObject.ApplicationExit(False)
       
    """
    def storyElev(self):
        NumberResults = 0
        Stories = []
        LoadCases = []
        StepTypes = []
        StepNums = []
        Directions = []
        Drifts = []
        Labels = []
        Xs = []
        Ys = []
        Zs = []

        [NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions, Drifts, Labels, Xs, Ys, Zs, self.ret] = \
            self.SapModel.Results.StoryDrifts(NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions,
                                                Drifts, Labels, Xs, Ys, Zs)
        
        elevs = []
        for Z in Zs:
            count = 0
            for elev in elevs:
                if elev == Z:
                    count += 1
            if count == 0:
                elevs.append(Z)
                
        elevs.reverse()
        return elevs        
        
    """

        
        
    #get all load combinaions
    def ComboName(self):
        NumberCombo = 0
        ComboNames = []
        [NumberCombo, ComboNames,ret] = self.SapModel.RespCombo.GetNameList(NumberCombo, ComboNames)
        return ComboNames
        
    #get all loads
    def LoadName(self):
        NumberCombo = 0
        ComboNames = []
        [NumberCombo, ComboNames,ret] = self.SapModel.LoadCases.GetNameList(NumberCombo, ComboNames)
        return ComboNames
        
    """
    def SelectedComcoStoryDriftResults(self, Combo):
        self.StoryDrifts = []
        self.ret = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        self.ret = self.SapModel.Results.Setup.SetComboSelectedForOutput(Combo)

        # initialize drift results
        NumberResults = 0
        Stories = []
        LoadCases = []
        StepTypes = []
        StepNums = []
        Directions = []
        Drifts = []
        Labels = []
        Xs = []
        Ys = []
        Zs = []

        [NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions, Drifts, Labels, Xs, Ys, Zs, self.ret] = \
            self.SapModel.Results.StoryDrifts(NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions,
                                                Drifts, Labels, Xs, Ys, Zs)
            # append all drift results to storydrifts list
        for i in range(0, NumberResults):
            self.StoryDrifts.append((Labels[i], Stories[i], LoadCases[i], Directions[i], 1000*Drifts[i], Xs[i], Ys[i], Zs[i]))

        # set up pandas data frame and sort by drift column
        labels = ["Labels",'Story', 'Combo', 'Direction', '1000*Drift',"Xs", "Ys", "Zs"]
        df =    pd.DataFrame.from_records(self.StoryDrifts, columns=labels)
        df['1000*Drift'] = df['1000*Drift'].round(4)
        dfDrop = df.sort_values('1000*Drift').drop_duplicates(subset=['Story', 'Direction'], keep='last')
        print(dfDrop)
        return dfDrop
    """
    
    """
    def SelectedLoadStoryDriftResults(self, Combo):
        self.StoryDrifts = []
        self.ret = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
        self.ret = self.SapModel.Results.Setup.SetCaseSelectedForOutput(Combo)

        # initialize drift results
        NumberResults = 0
        Stories = []
        LoadCases = []
        StepTypes = []
        StepNums = []
        Directions = []
        Drifts = []
        Labels = []
        Xs = []
        Ys = []
        Zs = []


        self.SapModel.Results.StoryDrifts(NumberResults, Stories, LoadCases, StepTypes, StepNums, Directions,
                                                Drifts, Labels, Xs, Ys, Zs)
            # append all drift results to storydrifts list
        for i in range(0, NumberResults):
            self.StoryDrifts.append((Labels[i], Stories[i], LoadCases[i], Directions[i], 1000*Drifts[i], Xs[i], Ys[i], Zs[i]))

        # set up pandas data frame and sort by drift column
        labels = ["Labels",'Story', 'Combo', 'Direction', '1000*Drift',"Xs", "Ys", "Zs"]
        df =    pd.DataFrame.from_records(self.StoryDrifts, columns=labels)
        df['1000*Drift'] = df['1000*Drift'].round(4)
        dfDrop = df.sort_values('1000*Drift').drop_duplicates(subset=['Story', 'Direction'], keep='last')
        print(dfDrop)
        return dfDrop
    """
    
    def storyDispForCombo(self, combo):
        self.ret = self.SapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay([combo])
        try:
            TableKey = 'Diaphragm Max Over Avg Drifts'
            FieldKeyList = []
            GroupName = str()
            TableVersion = 0
            FieldsKeysIncluded =[]
            NumberRecords = 0
            TableData = [] #table of Diaphragm Max Over Avg Drifts
            [GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret] = self.SapModel.DatabaseTables.GetTableForDisplayArray(TableKey, FieldKeyList,GroupName,TableVersion, FieldsKeysIncluded, NumberRecords, TableData)
            tableDrift = []
            if TableData[3] == "Max" or TableData[3] == "Min":
                for i in range(0, (len(TableData)), 12):
                    tableDrift.append([TableData[i],TableData[i+4][-1],(1000*float(TableData[i+5]))])
            else:
                for i in range(0, (len(TableData)), 11):
                    tableDrift.append([TableData[i],TableData[i+3][-1],(1000*float(TableData[i+4]))])                
            
            TableKey = 'Diaphragm Center Of Mass Displacements'
            FieldKeyList = []
            GroupName = str()
            TableVersion = 0
            FieldsKeysIncluded =[]
            NumberRecords = 0
            TableData = [] #table of Diaphragm Center Of mass Displacements
            [GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret] = self.SapModel.DatabaseTables.GetTableForDisplayArray(TableKey, FieldKeyList,GroupName,TableVersion, FieldsKeysIncluded, NumberRecords, TableData)
            tableDisplacement = []
            if TableData[4] == "Max" or TableData[4] == "Min":
                for i in range(0, (len(TableData)), 12):
                    tableDisplacement.append([TableData[i],(25.4*float(TableData[i+5])),(25.4*float(TableData[i+6]))]) #change inches to millimeters
            else:
                for i in range(0, (len(TableData)), 11):
                    tableDisplacement.append([TableData[i],(25.4*float(TableData[i+4])),(25.4*float(TableData[i+5]))]) #change inches to millimeters
 
                
            NumberStories = 0
            StoryNames = []
            StoryHeights = []
            StoryElevations = []
            IsMasterStory = []
            SimilarToStory = []
            SpliceAbove = []
            SpliceHeight = []
            ret = self.SapModel.Story.GetStories(NumberStories, StoryNames, StoryHeights, StoryElevations, IsMasterStory, SimilarToStory, SpliceAbove, SpliceHeight)
            storiesName = ret[1]
            StoryElevations = ret[2]
            
            
            labels = ['Story', "elevation", '1000DispX', '1000DispY', "1000DriftX", "1000DriftY"] #lable of data Frame
            table = []         
            k = 0
            # merg Diplacement list and Drift list 
            for story in storiesName:
                elev = StoryElevations[k]
                disX = 0
                disY = 0
                DriftX = 0
                DriftY = 0
                for dis in tableDisplacement:
                    if dis[0] == story:
                        dX = dis[1]
                        dY = dis[2]
                        if abs(dX) > abs(disX):
                            disX = dX
                        if abs(dY) > abs(disY):
                            disY = dY
                        
                for dif in tableDrift:
                    if dif[0] == story:
                        if dif[1] == "X":
                            DriftX = dif[2]
                        if dif[1]  == "Y":
                            DriftY = dif[2]
                table.append([story, elev,  disX, disY, DriftX, DriftY])
                k += 1
                
            df = pd.DataFrame.from_records(table, columns=labels).sort_index() #create dataframe
            return df
        except Exception as e:
            self.ConnectionsTATUS = False
            return
            
        """
            ret = self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
            ret = self.SapModel.Results.Setup.SetComboSelectedForOutput(combo)
            
            # initialize joint drift results
            NumberResults = 0
            Stories = []
            LoadCases = []
            Label  = ''
            Names = ''
            StepType = []
            StepNum = []
            # Directions = []
            DispX = []
            DispY = []
            DriftX = []
            DriftY = []

            [NumberResults, Stories, Label, Names, LoadCases, StepType, StepNum, DispX, DispY, DriftX, DriftY, ret] = \
                self.SapModel.Results.JointDrifts(NumberResults, Stories, Label, Names, LoadCases, StepType, StepNum,
                                                DispX, DispY, DriftX, DriftY)

            # append all displacement results to jointdrift list
            for i in range(0, NumberResults):
                self.JointDisplacements.append((Label[i], Stories[i], LoadCases[i], 1000*DispX[i], 1000*DispY[i], 1000*DriftX[i], 1000*DriftY[i]))

            # set up pandas data frame and sort by drift column
            jlabels = ['label', 'Story', 'Combo', '1000DispX', '1000DispY', "1000DriftX", "1000DriftY"]
            jdf = pd.DataFrame.from_records(self.JointDisplacements, columns=jlabels)
            return jdf 
    
    def StoryLevelS(self):
        NumberStories = 0
        StoryNames = []
        StoryHeights = []
        StoryElevations = []
        IsMasterStory = []
        SimilarToStory = []
        SpliceAbove = []
        SpliceHeight = []
        
        ret = self.SapModel.Story.GetStories(NumberStories, StoryNames, StoryHeights, StoryElevations, IsMasterStory, SimilarToStory, SpliceAbove, SpliceHeight)
        return StoryElevations
            """