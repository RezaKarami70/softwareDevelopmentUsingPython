import json
import pytest
import os
import Controller as controller
import Model as model
from unittest.mock import MagicMock



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
    pathFile = open("path.json")
    pathDict = json.load(pathFile)
    assert data == pathDict

