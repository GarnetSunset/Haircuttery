from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from Harvard import Excel2CSV
from IPython.display import HTML
from os.path import join, dirname, abspath
import csv
import glob
import itertools
import numpy as np
import os
import re
import requests
import sys
import threading
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

def loading(): 
   for s in itertools.cycle(['|','/','-','\\']):
      if done:
         break
      sys.stdout.write('\rloading ' + s)
      sys.stdout.flush()
      time.sleep(0.1)
      
done = False
      
worksheet = book.add_sheet("Results", cell_overwrite_ok=True)

url = "http://whoscall.in/1/*/"
headers = {'User-Agent': 'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}

file_paths = sys.argv[1:]
draganddrop = ''.join(file_paths)

response = requests.get(url, headers=headers)
content = BeautifulSoup(response.content, "lxml")

if draganddrop == "":
   fileName = raw_input("\nInput the file with extension\n>")
   stopPoint = fileName.index('.')
   prepRev = fileName[stopPoint:]
   csvTest = prepRev   
else:
   fileName = draganddrop
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
#http://whoscall.in/1/**********/ is the layout#

xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

website = raw_input("Input 1 for whoscall.in results, input 2 for BBB\n>")
numFormat = raw_input("Which format?\n2 for Simple, 1 for Complex, 0 for Normal\n>")

g = threading.Thread(target=loading)
g.start()

if(website =="1"):
   stopPoint = fileName.index('.')
   prepRev = fileName[0:stopPoint]
   totalName = prepRev + "_rev_who.xlsx"   
   workbook = xlsxwriter.Workbook(totalName)
   worksheet = workbook.add_worksheet()   
   worksheet.write(0,0, "Telephone Number")
   worksheet.write(0,1, "# of Messages")
   worksheet.write(0,2, "Does it Appear?")
   worksheet.write(0,3, "Number of Scammers")
   worksheet.write(0,4, "Number of Spammers")
   worksheet.write(0,5, "Number of Debt Collectors")
   worksheet.write(0,6, "Number of Hospital")
   worksheet.write(0,7, "Sentiment")

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
   
   if(numFormat == "0"):
      firstStart = cell_obj_str.index('-')-3
      firstEnd = firstStart + 3   
      secondStart = cell_obj_str.index('-')+1
      secondEnd = secondStart + 3
      thirdStart = cell_obj_str.index('-')+5
      thirdEnd = thirdStart + 4      
      teleWho = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      teleBBB = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])   
   
   if(numFormat == "1"):      
      firstStart = cell_obj_str.index('(')+1
      firstEnd = firstStart + 3   
      secondStart = cell_obj_str.index(' ')+1
      secondEnd = secondStart + 3
      thirdStart = cell_obj_str.index('-')+1
      thirdEnd = thirdStart + 4
      teleWho = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      teleBBB = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      
   if(numFormat == "2"):
      teleWho = (cell_obj_str[8:11] + cell_obj_str[11:14] + cell_obj_str[14:18])
      teleBBB = (cell_obj_str[8:11] + cell_obj_str[11:14] + cell_obj_str[14:18])
      
   #print('(%s) %s' % (idx, teleWho))
   perm = teleWho
   site800 = teleWho
   worksheet.write(idx+1, 0, perm)
   
   if(website == "1"):  
      reqInput = "http://whoscall.in/1/%s/" % (teleWho)
      urlfile = urllib2.Request(reqInput)
      #print (reqInput)#
      time.sleep(1)
      requestRec = requests.get(reqInput)
      soup = BeautifulSoup(requestRec.content, "lxml")
      noMatch = soup.find(text=re.compile(r"no reports yet on the phone number"))
      #print(requestRec.content)###only if needed#
      type(noMatch) is str
      if noMatch is None:
         worksheet.write(idx+1, 2, "Got a hit")
         howMany = soup.find_all('img',{'src':'/default-avatar.gif'})
         howManyAreThere = len(howMany)
         worksheet.write(idx+1,1,howManyAreThere)
         #print (howManyAreThere)
         scamNum = [ div for div in soup.find_all('div', {'style':'font-size:14px; margin:10px; overflow:hidden'}) if 'scam' in div.text.lower() or 'Scam' in div.text.lower() or 'scams' in div.text.lower() ]
         scamCount = len(scamNum)
         spamNum = [ div for div in soup.find_all('div', {'style':'font-size:14px; margin:10px; overflow:hidden'}) if 'spam' in div.text.lower() or 'Spam' in div.text.lower() or 'spams' in div.text.lower() ]
         spamCount = len(spamNum)     
         debtNum = [ div for div in soup.find_all('div', {'style':'font-size:14px; margin:10px; overflow:hidden'}) if 'debt' in div.text.lower() or 'Debt' in div.text.lower() or 'credit' in div.text.lower() ]
         debtCount = len(debtNum)
         hospitalNum = [ div for div in soup.find_all('div', {'style':'font-size:14px; margin:10px; overflow:hidden'}) if 'hospital' in div.text.lower() or 'Hospital' in div.text.lower() ]
         hospitalCount = len(hospitalNum)
         if hospitalCount > 0:
            hospitalCount+9999
         searchTerms = {'Scam':scamCount,'Spam':spamCount,'Debt Collector':debtCount,'Hospital':hospitalCount}
         sentiment = max(searchTerms, key=searchTerms.get) 
         worksheet.write(idx+1,3,scamCount)
         worksheet.write(idx+1,4,spamCount)
         worksheet.write(idx+1,5,debtCount)
         worksheet.write(idx+1,6,hospitalCount)
         worksheet.write(idx+1,7,sentiment)
         if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0:
            worksheet.write(idx+1,7,"No Entries Detected")
        
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

prepRev = prepRev + '_temp.csv'

Excel2CSV(totalName, "Sheet1", prepRev)

done = True

if delMe == 1:
   os.remove(deleteFile)
   os.remove(prepRev)
   print("Temp File Cleaned!\n")

ding = "Ding! Job Done!"
uni = unicode( ding, "utf-8")
bytesNow = uni.encode( "utf-8" )
print (ding)