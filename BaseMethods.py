import xlrd, xlwt
from xlutils.copy import copy
from PyQt5.QtCore import *

def GetValue(text):
    text = text.upper()
    r = 0
    l = 0
    while (l < len(text)):
        if not text[l].isalpha():
            print('Can not handle non-alphabets character.')
            return -1
        r += (ord(text[l]) - ord('A') + 1) * 26**(len(text)-l-1)
        l += 1
    return r - 1

def GetDistance(text1, text2):
    v1 = GetValue(text1)
    v2 = GetValue(text2)
    if v1 == -1 or v2 == -1:
        print('This method can only handle characters.')
    return v2 - v1

def IsNumber(value):
    try:
        float(value)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(value)
        return True
    except (TypeError, ValueError):
        pass
 
    return False

def GetRdSheet(filePath, sheetIndex):
    try:
        rb = xlrd.open_workbook(filePath)
        return rb.sheet_by_index(sheetIndex)
    except Exception as e:
        print(e)
        raise Exception('打开文件[{}], sheet[{}]时错误'.format(filePath, sheetIndex))

def GetWtSheet(filePath, sheetIndex):
    try:
        rb = xlrd.open_workbook(filePath)
        wb = copy(rb)
        return wb, wb.get_sheet(sheetIndex), rb.sheet_by_index(sheetIndex)
    except Exception as e:
        print(e)
        raise Exception('打开文件[{}], sheet[{}]时错误'.format(filePath, sheetIndex))

def FindMatchedCodeIndex(sheet, code, isColumn, index):
    if isColumn:
        values = sheet.col_values(index)
    else:
        values = sheet.row_values(index)
    i = 0
    for item in values:
        if item == code:
            return i
        i += 1
    if isColumn:
        raise Exception('在列[{}]中找不到码[{}]'.format(index, code))
    else:
        raise Exception('在行[{}]中找不到码[{}]'.format(index, code))

def FindMatchedCodeIndexInList(values, code):
    i = 0
    for item in values:
        if IsNumber(item):
            if item == code:
                return i
        i += 1
    raise Exception('找不到码[{}]'.format(code))

def ChangeCellValue(sheet, row, col, value):
    sheet.write(row, col, value)

def FindAndInsert(fp1, sheetIdx1, isColumn1, codeidx1, numidx1, fp2, sheetIdx2, isColumn2, codeidx2, numidx2, resultName):
    sheetRD = GetRdSheet(fp1, sheetIdx1)
    wb, sheetWT, sheetWR = GetWtSheet(fp2, sheetIdx2)
    if isColumn1:
        codesRD = sheetRD.col_values(codeidx1)
        numsRD = sheetRD.col_values(numidx1)
    else:
        codesRD = sheetRD.row_values(codeidx1)
        numsRD = sheetRD.row_values(numidx1)
    
    if isColumn2:
        codesWT = sheetWR.col_values(codeidx2)
    else:
        codesWT = sheetWR.row_values(codeidx2)

    # print('codesRD:{}'.format(codesRD))
    # print('numsRD:{}'.format(numsRD))
    # print('codesWT:{}'.format(codesWT))

    i = -1
    for item in codesRD:
        i += 1
        if IsNumber(item):
            try:
                index = FindMatchedCodeIndexInList(codesWT, item)
                print('code[{}] -> index[{}]'.format(item, index))
                if isColumn2:
                    ChangeCellValue(sheetWT, index, numidx2, numsRD[i])
                else:
                    ChangeCellValue(sheetWT, numidx2, index, numsRD[i])
            except:
                continue    
    wb.save(resultName)

def GetSheets(filePath):
    rb = xlrd.open_workbook(filePath)
    return rb.sheet_names()

if __name__ == '__main__':
    print(GetValue('A'))
    print(GetValue('AA'))
    print(GetValue('BA'))
    print(GetValue('cc'))
    print(GetValue('aaa'))
    print(GetDistance('A', 'aa'))