from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from flask import Flask, jsonify, make_response, abort, url_for, request
from googleapiclient import sample_tools
from Harvard import Excel2CSV
from IPython.display import HTML
from os.path import join, dirname, abspath
import csv
import glob
import numpy as np
import os
import re
import requests
import time
import urllib2
import xlrd
import xlsxwriter
import xlwt

delMe = 0
w=1
x=2
y=3
z=4

#list800=[cell_obj_str,total_messages,category,last_three]
#listBBB=[cell_obj_str,category,accredited]

book = xlwt.Workbook(encoding="utf-8")

worksheet = book.add_sheet("Results", cell_overwrite_ok=True)

url = "http://800notes.com/Phone.aspx/*"
headers = {'User-Agent': 'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}

response = requests.get(url, headers=headers)
content = BeautifulSoup(response.content, "lxml")
fileName = raw_input("\nInput the file with extension\n>")
stopPoint = fileName.index('.')
prepRev = fileName[stopPoint:]
csvTest = prepRev

if csvTest == ".csv":
   excelFile = xlsxwriter.Workbook(fileName + '.xlsx')
   worksheet = excelFile.add_worksheet()
   with open(fileName,'rb') as f:
      content = csv.reader(f)
      for index_col, data_in_col in enumerate(content):
            for index_row, data_in_cell in enumerate(data_in_col):
                  worksheet.write(index_col,index_row,data_in_cell)
   excelFile.close()
   fileName = (fileName + '.xlsx')
   delMe = 1
   print("Temporary Convert to xlsx done.\n")

deleteFile = fileName

fname = join(dirname(abspath('__file__')), '%s' % fileName)
#http://800notes.com/Phone.aspx/*-***-***-**** is the layout#

xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

website = raw_input("Input 1 for 800Notes Google results, input 2 for BBB\n>")

if(website =="1"):
   stopPoint = fileName.index('.')
   prepRev = fileName[0:stopPoint]
   totalName = prepRev + "_rev_800.xlsx"   
   workbook = xlsxwriter.Workbook(totalName)
   worksheet = workbook.add_worksheet()   
   worksheet.write(0,0, "Telephone Number")
   worksheet.write(0,1, "# of Messages")
   worksheet.write(0,2, "Category")
   worksheet.write(0,3, "Last 3 Messages")

if(website == "2"):
   stopPoint = fileName.index('.')
   prepRev = fileName[0:stopPoint]
   totalName = prepRev + "_rev_BBB.xlsx"
   workbook = xlsxwriter.Workbook(totalName)
   worksheet = workbook.add_worksheet()
   worksheet.write(0,0, "Telephone Number")
   worksheet.write(0,1, "Acreditted")

worksheet.set_column('A:A',13)

col = xl_sheet.col_slice(0,1,10101010)
from xlrd.sheet import ctype_text
#print('(Column #) type:value')
for idx, cell_obj in enumerate(col):
   cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')  
   cell_obj_str = str(cell_obj)
   
   tele800 = (cell_obj_str[8:11] + "-" + cell_obj_str[11:14] + "-" + cell_obj_str[14:18])
   teleBBB = ("%28" + cell_obj_str[8:11] + "%29+" + cell_obj_str[11:14] + "-" + cell_obj_str[14:18])
   print('(%s) %s' % (idx, tele800))
   perm = tele800
   site800 = tele800 + " site:800notes.com"
   worksheet.write(idx+1, 0, perm)
    
   if(website == "1"):  
      reqInput = "http://www.google.com/search?q=%s+site:800notes.com&num=100&hl=en&start=0" % (tele800)
      urlfile = urllib2.Request(reqInput)
      page = urlfile.read()
      soup = BeautifulSoup(page)
      print (reqInput)
      time.sleep(10)
      requestRec = requests.get(reqInput)
      soup = BeautifulSoup(requestRec.content, "lxml")
      print(requestRec.content)
      noMatch = soup.body.findAll(text='did not match any documents.')
      print(noMatch)
      if len(noMatch)==0:
         worksheet.write(idx+1, 2, "Got a hit")     
        
   if(website == "2"):
      reqInput = ('https://www.bbb.org/search/?splashPage=true&type=name&input='+ teleBBB +'&location=&tobid=&filter=business&radius=&country=USA%2CCAN&language=en&codeType=YPPA')
      print (reqInput)
      time.sleep(1)
      requestRec = requests.get(reqInput)
      soup = BeautifulSoup(requestRec.content,"lxml")
      divTags = soup.find('class')
      Badge = soup.find_all('img',{'class':'badge-accredited'})
      Hit = soup.find_all('td',{'class':'accredited'})
      if len(Hit)!=0:
         worksheet.write(idx+1,1,"Got a Hit")      
      if len(Badge)!=0:
         worksheet.write(idx+1,1,"Is Accredited")
         
workbook.close()

Excel2CSV(totalName, "Sheet1", prepRev + '.csv')

if delMe == 1:
   os.remove(deleteFile)
   print("Temp File Cleaned!\n")

print("Ding! Job Done! ᕕ( ᐛ )ᕗ")
