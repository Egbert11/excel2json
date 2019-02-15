#!/usr/bin/env python
#-*- coding: utf-8 -*-

# function: 将Excel文件转换为json文件
# create_time: 2019/2/15

import sys
import SheetManager

def export_json():
    file_path = sys.argv[1]
    file_name = file_path[:file_path.rfind(".")]
    SheetManager.addWorkBook(file_path)
    sheetNameList = SheetManager.getSheetNameList()

    for sheet_name in sheetNameList:
        sheetJSON = SheetManager.exportJSON(sheet_name)

        f = file(file_name+sheet_name+'.json', 'w')
        f.write(sheetJSON.encode('UTF-8'))
        f.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print "usage: python excel_2_json.py xxx.xlsx"
        sys.exit(1)
    export_json()

