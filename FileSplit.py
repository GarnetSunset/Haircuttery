from Harvard import enumColumn
from os.path import join, dirname, abspath
from shutil import copyfile
from shutil import move
from xlrd.sheet import ctype_text
import csv
import os
import sys
import time
import xlrd
import xlsxwriter
import xlwt

dragNDrop = ''.join(sys.argv[1:])
beginGrab = 1
counting = 0
delMe = 0
thousands = 0

if dragNDrop == "":
    fileName = raw_input("\nInput the file with extension\n>")
else:
    fileOnly = dragNDrop.rfind('\\') + 1
    fileName = dragNDrop[fileOnly:]
stopPoint = fileName.index('.')
prepRev = fileName[stopPoint:]
preName = fileName[:stopPoint]
if(fileName.rfind('\\') != None):
    postSlash = fileName.rfind('\\') + 1
    preName = fileName[postSlash:stopPoint]
reload(sys)
sys.setdefaultencoding('utf')
# Is the extension CSV? If so we'll convert it to xlsx.
if prepRev == ".csv":
    totalName = preName + '.xlsx'
    excelFile = xlsxwriter.Workbook(totalName)
    worksheet = excelFile.add_worksheet()
    enumColumn(fileName, worksheet)
    excelFile.close()
    fname = join(dirname(abspath('__file__')), '%s' % totalName)
    print('Temporary Convert to xlsx done.\n')
else:
    totalName = preName + prepRev
fname = join(dirname(abspath('__file__')), '%s' % totalName)
stopPoint = fileName.index('.')
prepRev = fileName[0:stopPoint]
xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])
book = xlwt.Workbook(encoding="utf-8")
worksheet = book.add_sheet("Results", cell_overwrite_ok=True)
workbook = xlrd.open_workbook(fileName)
for sheet in workbook.sheets():
    for row in range(sheet.nrows):
        row = int(row)
if(int(row) > 1000):
    subDivide = int(row) / 1000
    while(thousands != subDivide + 1):
        thousands = thousands + 1
        counting = 0
        totalName = preName + "_" + str(thousands) + ".xlsx"
        totalFile = xlsxwriter.Workbook(str(totalName))
        totalFile.close()
        totalFile = xlsxwriter.Workbook(str(totalName))
        worksheet = totalFile.add_worksheet()
        worksheet.write(0, 0, "Telephone Number")
        with open(totalName, 'rb') as f:
            col = xl_sheet.col_slice(0, beginGrab, 10101010)
            for idx, cell_obj in enumerate(col):
                counting = counting + 1
                if(counting == 1000):
                    break
                cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
                cell_obj_str = str(cell_obj)
                telePhone = (cell_obj_str[7:19])
                telePhone = telePhone.replace("'", "")
                worksheet.write(counting, 0, telePhone)
        beginGrab += 1000
        totalFile.close()
        totalFile = None
        if not os.path.exists("WorkingDir"):
            os.makedirs("WorkingDir")
        if not os.path.exists("WorkingDir/" + preName):
            os.makedirs("WorkingDir/" + preName)
        if(totalName.rfind('\\') != None):
            postSlash = fileName.rfind('\\') + 1
            folderName = totalName[postSlash:]
            move(folderName, "WorkingDir/" + preName + '/' + folderName)
        else:
            move(totalName, "WorkingDir/" + preName + '/' + totalName)
    copyfile(fileName, "WorkingDir/" + preName + '/' + fileName)
    copyfile('HairCut.py', "WorkingDir/" + preName + '/HairCut.py')
    copyfile('Harvard.py', "WorkingDir/" + preName + '/Harvard.py')
    tempFile = open('tempFile.log', 'w')
    tempFile.write(preName)
    tempFile.close()
    if os.path.isfile('chrome.ini'):
        copyfile('chrome.ini', "WorkingDir/" + preName + '/chrome.ini')
    if os.path.isfile('chromedriver.exe'):
        copyfile('chromedriver.exe', "WorkingDir/" +
                 preName + '/chromedriver.exe')

    if delMe == 1:
        os.remove(fileName)
        print("Temp File Cleaned!\n")

else:
    if not os.path.exists("WorkingDir"):
        os.makedirs("WorkingDir")
    if not os.path.exists("WorkingDir/" + preName):
        os.makedirs("WorkingDir/" + preName)
    copyfile(fileName, "WorkingDir/" + preName + "/" + fileName)
    copyfile('HairCut.py', "WorkingDir/" + preName + '/HairCut.py')
    copyfile('Harvard.py', "WorkingDir/" + preName + '/Harvard.py')
    tempFile = open('tempSmall.log', 'w')
    tempFile.write(preName)
    tempFile.close()
    if os.path.isfile('chrome.ini'):
        copyfile('chrome.ini', "WorkingDir/" + preName + '/chrome.ini')
    if os.path.isfile('chromedriver.exe'):
        copyfile('chromedriver.exe', "WorkingDir/" +
                 preName + '/chromedriver.exe')
    sys.exit()
print('Ding! Job Done!')
