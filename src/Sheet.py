# encoding: utf-8

import xlrd
import json
import time
from xlrd import sheet

from xlrd import XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, XL_CELL_DATE, XL_CELL_BOOLEAN, XL_CELL_ERROR, \
    XL_CELL_BLANK

arrPrefix = "arr"
arrPrefixLen = len(arrPrefix)

class ValType:
    int = "int"
    float = "float"
    string = "string"
    dict = "dict"
    date = "date"
    day = "day"
    # 数组以arr开头，如arrint, arrfloat

class Sheet:
    def __init__(self, sh):
        self.sh = sh
        self.name = sh.name
        #解析的数据
        self.python_obj = {}

        self.__findRow()

        self.__parseField()


    # 查找数据起始行数，数据终止行数
    def __findRow(self):
        self.dataStartRow = 0
        self.dataEndRow = -1

        for row in range(self.sh.nrows):
            if self.sh.cell(row, 0).ctype == XL_CELL_EMPTY:
                print row
                self.dataEndRow = row
                break
        if self.dataEndRow == -1:
            self.dataEndRow = self.sh.nrows


    def __check(self, row):
        for col in range(0, 4):
            cell = self.sh.cell(row, col)
            if cell.ctype == XL_CELL_EMPTY:
                print "invalid format row:%s col:%s" % (row, col)
                exit(1)

    # 解析字段属性
    def __parseField(self):
        if self.dataStartRow >= self.dataEndRow:
            print "empty excel, can't not output json file"
            exit(1)
        for row in range(self.dataStartRow, self.dataEndRow):
            self.__check(row)
            desc = self.sh.cell(row, 0).value
            val_type = self.sh.cell(row, 1).value
            if val_type not in ValType.__dict__ and not val_type.startswith(arrPrefix):
                print "invalid val_type for row:%s" % (row)
                exit(1)
            key = self.sh.cell(row, 2).value
            val = self.sh.cell(row, 3).value
            print "key:{},val:{}".format(key, val)
            if not val_type.startswith(arrPrefix):
                val = self.__parseRouter(val, val_type)
            else:
                val = self.__parseArray(row, 3, val_type[arrPrefixLen:])

            self.python_obj[key] = val

    def __parseRouter(self, val, type):
        d = {
            ValType.int: self.__parseInt,
            ValType.float: self.__parseFloat,
            ValType.string: self.__parseString,
            ValType.dict: self.__parseDict,
            ValType.date: self.__parseDate,
            ValType.day: self.__parseDay,
        }
        if type not in d:
            print "unknown type:%s" % type
            exit(1)
        return d.get(type)(val)

    # 优先转成int，次之float，最后是string
    def __parseIntFloatString(self, val):
        if not val:
            return ""
        try:
            num = int(val)
        except ValueError:
            pass
        else:
            return num
        try:
            num = float(val)
        except ValueError:
            pass
        else:
            return num
        return val

    # 转换字符串为int
    def __parseInt(self, val):
        num = 0
        try:
            num = int(val)
        except ValueError:
            pass
        return num

    # 转换字符串为float
    def __parseFloat(self, val):
        num = 0
        try:
            num = float(val)
        except ValueError:
            pass
        return num

    # 转换字符串为string
    def __parseString(self, val):
        if not val:
            return ""
        res = ""
        try:
            res = str(val)
        except BaseException:
            pass
        return res

    # 转换字符串为dict
    def __parseDict(self, str):
        dict = {}
        list = str.split(',')
        for i in range(len(list)):
            kv = list[i].split(':')
            key = kv[0]
            value = kv[1]
            dict[key] = self.__parseIntFloatString(value)

        return dict

    # 将date转为时间戳
    def __parseDate(self, val):
        res = 0
        date = self.__parseFloat(val)
        try:
            timetuple = xlrd.xldate_as_tuple(date, 0)  # 0: 1900-based, 1: 1904-based.
            timetuple += (0, 0, 0)
            res = int(time.mktime(timetuple))
        except BaseException:
            pass
        return res

    # 将day转为string，如2019-02-15
    def __parseDay(self, val):
        res = "1900-01-01"
        date = self.__parseFloat(val)
        try:
            timetuple = xlrd.xldate_as_tuple(date, 0)  # 0: 1900-based, 1: 1904-based.
            res = "{}-{}-{}".format(timetuple[0],timetuple[1],timetuple[2])
        except BaseException:
            pass
        return res

    # 转换成数组
    def __parseArray(self, row, beginCol, type):
        arr = []
        while beginCol < self.sh.ncols:
            cell = self.sh.cell(row, beginCol)
            if cell.ctype == XL_CELL_EMPTY:
                break
            val = cell.value
            val = self.__parseRouter(val, type)
            arr.append(val)
            print val
            beginCol += 1
        return arr


    def toJSON(self):
        print self.python_obj
        json_obj = json.dumps(self.python_obj, sort_keys=True, indent=2, ensure_ascii=False)
        return json_obj


def openSheet(sh):
    return Sheet(sh)