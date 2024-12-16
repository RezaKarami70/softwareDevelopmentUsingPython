import os, comtypes.client, json
import pandas as pd
import ctypes
from ctypes import wintypes


class Model():
    
    ConnectionsTATUS = False

    def runModel(self):
        self.pathFile = open("path.json")
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
        

        
        
        
    def ComboName(self):
        
        NumberCombo = 0
        ComboNames = []
        [NumberCombo, ComboNames,ret] = self.SapModel.RespCombo.GetNameList(NumberCombo, ComboNames)
        return ComboNames
        

    def LoadName(self):
        
        NumberCombo = 0
        ComboNames = []
        [NumberCombo, ComboNames,ret] = self.SapModel.LoadCases.GetNameList(NumberCombo, ComboNames)
        return ComboNames
        
       
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
    
    def storyDispForCombo(self, combo):
        # returns dataframe of torsion results for drift combinations
        self.JointDisplacements = []
        try:
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
        except Exception as e:
            self.ConnectionsTATUS = False
            return
    
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