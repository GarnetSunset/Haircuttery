# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/#
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
