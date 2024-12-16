import json
import pytest
import os
import Controller as controller



def test_load_config_valid():
    pathFile = open("path.json")
    pathDict = json.load(pathFile)
    con = controller.Controller()
    result = con.loadJson()
    assert result == pathDict


def test_save_config_valid():
    data = {"key": "value"}
    con = controller.Controller()
    con.pathDict = data
    con.saveJson()
    print(data)
    pathFile = open("path.json")
    pathDict = json.load(pathFile)
    print(str(pathDict))
