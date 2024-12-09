import json


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
        except (FileNotFoundError, json.decoder.JSONDecodeError):
            a_dict = {"etabs_path": "Import path to etabs", "model_path": "Import path to model", "model_name": "Import model name",
                    "folder_path": "Import path to saving folder", "previously_connected_to": ""}
            with open('path.json', 'w') as outfile:
                json.dump(a_dict, outfile)
        else:
            break
        
        
    import View

    view = View.View()

    view.fillPathWindow()
    view.pathWindow.exitButton.pressed.connect(view.pathWindowExitButtonAction)
    view.pathWindow.closeButton.pressed.connect(view.pathWindowCloseButtonAction)
    view.pathWindow.runButton.pressed.connect(view.pathWindowRunButtonAction)
    view.mainWindow.driftControlButton.pressed.connect(view.mainWindowDriftControlButtonAction)
    view.mainWindow.pathButton.pressed.connect(view.mainWindowpathButtonAction)
    view.mainWindow.closeButton.pressed.connect(view.mainWindowCloseButtenAction)
    view.mainWindow.connectButton.pressed.connect(view.mainWindowConnectButtonAction)
    view.driftControlWindow.closeButton.pressed.connect(view.driftControlWindowCloseACtions)
    view.app.exec()