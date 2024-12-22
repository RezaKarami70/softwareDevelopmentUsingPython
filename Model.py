import os, comtypes.client, json
import pandas as pd


class Model():

    def getModel(self):
        """cacth an etabs model and run analysis if exist run analysis if no etabs model is run in the os"""
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
        """ 
        cacth an etabs model and run analysis if exist or open etabs model whit run Etabs.exe ath and
        run analysis if no etabs model is run in the os
        """   
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

        
    def ComboName(self):
        """
        get all load combinaions from etabs API and return them as a list
        """
        NumberCombo = 0
        ComboNames = []
        [NumberCombo, ComboNames,ret] = self.SapModel.RespCombo.GetNameList(NumberCombo, ComboNames)
        return ComboNames
    
    def storyDispForCombo(self, combo):
        """
        sending a load combination to Etabs API and get 3 tables
        1- 'Diaphragm Max Over Avg Drifts' for getting the maximum drift of each story
        2- 'Diaphragm Center Of Mass Displacements' for Getting the center of mass displacement of each story
        3- Getting storys from Story interface to know hight of each story
        
        input: Model as self and a load combination as a string from the selleced load combination in the list of 
        drift control window

        return a pandas Data Frame including  'Story Name', "elevation", 'DispX mm', 'DispY mm', "DriftX mm", "DriftY mm"
        if any excetions , return False
        """
        self.SapModel.SetPresentUnits(6) #set the API Tables to Metric
        self.ret = self.SapModel.DatabaseTables.SetLoadCombinationsSelectedForDisplay([combo])
        try:
            table_key = 'Diaphragm Max Over Avg Drifts'
            field_key_list = 'Story', 'OutputCase', 'StepType', 'StepNumber', 'Item', 'Max Drift', 'Label'
            requested_field_keys, version, field_keys, num_rec, table, ret = self.SapModel.DatabaseTables.GetTableForDisplayArray(table_key, field_key_list, None)
            df_Dif = pd.DataFrame(columns=field_key_list)
            step = len(field_keys)
            for i, header in enumerate(field_keys):
                df_Dif[header] = table[i::step]
                
            TableKey = 'Diaphragm Center Of Mass Displacements'
            FieldKeyList = 'Story', 'Diaphragm', 'OutputCase', 'CaseType', 'UX', 'UY'
            GroupName = str()
            TableVersion = 0
            FieldsKeysIncluded =[]
            NumberRecords = 0
            TableData = [] #table of Diaphragm Center Of mass Displacements
            GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret = self.SapModel.DatabaseTables.GetTableForDisplayArray(TableKey, FieldKeyList,GroupName,TableVersion, FieldsKeysIncluded, NumberRecords, TableData)            
            df_Dis = pd.DataFrame(columns=FieldKeyList)
            step = len(FieldsKeysIncluded)
            for i, header in enumerate(FieldsKeysIncluded):
                df_Dis[header] = TableData[i::step]
            
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
            base = ["base",0.0,0.0,0.0,0.0,0.0]
            table.append(base)
            k = 1
            # merg Diplacement list and Drift list 
            for i in range(1 , len(storiesName), 1):
                story = storiesName[i]
                DirList = df_Dif[( df_Dif["Story"] == story)]["Item"].drop_duplicates().to_list()
                Xtext = ""
                Ytext = ""
                if len(DirList) == 2:
                    Xtext = DirList[0]
                    Ytext = DirList[1]
                if len(DirList) == 1:
                    if DirList[0][-1] == "X":
                        Xtext = DirList[0]
                    if DirList[0][-1] == "Y":
                        Ytext = DirList[0]
                elev = float(StoryElevations[k])
                disX = float(df_Dis[( df_Dis["Story"] == story)]["UX"].max())
                disY = float(df_Dis[( df_Dis["Story"] == story)]["UY"].max())
                dif_X = abs(float(df_Dif[( df_Dif["Story"] == story) & (df_Dif["Item"] == str(Xtext))]["Max Drift"].max()))
                dif_Y = abs(float(df_Dif[( df_Dif["Story"] == story) & (df_Dif["Item"] == str(Ytext))]["Max Drift"].max()))
                table.append([story, elev, disX*1000, disY*1000, dif_X*1000, dif_Y*1000])
                k += 1
            df = pd.DataFrame.from_records(table, columns=labels).sort_index() #create dataframe
            for l in labels:
                df[l] = df[l].fillna(0)
            return df
        except Exception as e:
            self.ConnectionsTATUS = False
            return
            