#https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/#
import csv
import xlwt
import xlrd

# Go through those CSV columns!


def enumColumn(fileName, worksheet):
    with open(fileName, 'rb') as f:
        content = csv.reader(f)
        for index_col, data_in_col in enumerate(content):
            for index_row, data_in_cell in enumerate(data_in_col):
                worksheet.write(index_col, index_row, data_in_cell)

# Convert Excel to CSV


def Excel2CSV(ExcelFile, SheetName, CSVFile):
    import csv
    import xlrd
    workbook = xlrd.open_workbook(ExcelFile)
    worksheet = workbook.sheet_by_name(SheetName)
    csvfile = open(CSVFile, 'wb')
    wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
    for rownum in xrange(worksheet.nrows):
        wr.writerow(
            list(x.encode('utf-8') if type(x) == type(u'') else x
                 for x in worksheet.row_values(rownum)))
    csvfile.close()
