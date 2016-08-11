#https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/#
import xlwt
import xlrd
import csv
from HTMLParser import HTMLParser

def construct_entries():
    entries = []
    for i in xrange(len(list)):
        entry = {}
        entry['id'] = i
        entry['area'] = get_area_code(list[i])
        entry['number'] = get_full_number(list[i])
        entry['comment'] = get_comment(list[i])
        entries.append(entry)
    return entries

def get_full_number(entry):
    return entry.find('a', {"class": "oos_previewTitle"}).getText()

def get_area_code(entry):
    full_number = get_full_number(entry)
    return full_number[:3]

def get_comment(entry):
    comment = {}
    comment['count'] = get_comment_number(entry)
    comment['content'] = get_comment_content(entry)
    return comment

def get_comment_number(entry):
    return entry.find('span', {"class": "postCount"}).getText()

def get_comment_content(entry):
    return entry.find('div', {"class": "oos_previewBody"}).getText()

def Excel2CSV(ExcelFile, SheetName, CSVFile):
    import xlrd
    import csv
    workbook = xlrd.open_workbook(ExcelFile)
    worksheet = workbook.sheet_by_name(SheetName)
    csvfile = open(CSVFile, 'wb')
    wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

    for rownum in xrange(worksheet.nrows):
        wr.writerow(
            list(x.encode('utf-8') if type(x) == type(u'') else x
                 for x in worksheet.row_values(rownum)))

    csvfile.close()