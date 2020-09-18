# -*- coding: utf-8 -*-
"""
    
    print_utility.py
   ----------------------
   
   This module contains functions for printint to a cell using xlwings.

"""

__author__ = "Steven M. Satterfield"

import constant
import xlwings as xw
from xlwings import constants
from xlwings.utils import rgb_to_int

wsEmittersRow = 2
wsModesRow = 2
wsGeneratorsRow = 2
wsPRISequencesRow = 2
wsFREQSequencesRow = 2


def writeCell(ws, wsRow, cellValue):
    ws.range(wsRow, 1).value = cellValue
    ws.range(wsRow, 1).api.VerticalAlignment = constants.VAlign.xlVAlignTop
    #ws.range(wsRow, 1).WrapText = True

    
def writeSpecificCell(ws, wsRow, wsCol, cellValue):
    ws.range(wsRow, wsCol).value = cellValue
    ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignCenter
    ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
    ws.range(wsRow, wsCol).WrapText = True

def writeValueCell(ws, wsRow, wsCol, cellValue):
    ws.range(wsRow, wsCol).value = cellValue
    try:
        ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignLeft
        ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
        ws.range(wsRow, wsCol).WrapText = True
    except Exception:
        print("exception was thrown")
        

def writeLabelCell(ws, wsRow, wsCol, cellValue):
    ws.range(wsRow, wsCol).value = cellValue
    try:
        ws.range(wsRow, wsCol).api.HorizontalAlignment = constants.HAlign.xlHAlignRight
        ws.range(wsRow, wsCol).api.VerticalAlignment = constants.VAlign.xlVAlignTop
        ws.range(wsRow, wsCol).WrapText = True
        if cellValue == constant.XL_MISSING_TEXT:
            ws.range(wsRow, wsCol).api.Font.Color = rgb_to_int((255,0,0))

    except Exception:
        print("exception was thrown")
