# encoding: utf-8

import xlrd
import json
import Sheet

sheetDict = {}
sheetNameList = []

def addWorkBook(filepath):
    wb = xlrd.open_workbook(filepath)

    for sheet_index in range(wb.nsheets):
        sh = wb.sheet_by_index(sheet_index)
        sheet = Sheet.openSheet(sh)
        addSheet(sheet)

def addSheet(sheet):
    sheetDict[sheet.name] = sheet
    sheetNameList.append(sheet.name)

def getSheet(name):
    return sheetDict[name]

def getSheetNameList():
    return sheetNameList

def exportJSON(name):
    return sheetDict[name].toJSON()

def isReferencedSheet(name):
    for sheetName in sheetDict:
        if name in sheetDict[sheetName].referenceSheets:
            return  True

    return False