# -*- coding: utf-8 -*-
from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from Harvard import Excel2CSV
from IPython.display import HTML
from os.path import join, dirname, abspath
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
w = 1
x = 2
y = 3
z = 4
countitup = 0
scamCount = 0
spamCount = 0
debtCount = 0
hospitalCount = 0

def loading(): 
   for s in itertools.cycle(['|','/','-','\\']):
      if done:
         break
      sys.stdout.write('\rloading ' + s)
      sys.stdout.flush()
      time.sleep(0.1)
      
done = False
url = "http://whoscall.in/1/*/"
book = xlwt.Workbook(encoding="utf-8")
headers = {'User-Agent': 'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}
worksheet = book.add_sheet("Results", cell_overwrite_ok=True)
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

website = raw_input("Input 1 for whoscall.in results, input 2 for BBB, input 3 for You Can't Call Us\n>") #, input 3 for 800notes
numFormat = raw_input("Which format?\n1 for xxx-xxx-xxxx, 2 for (xxx) xxx-xxxx, 3 for xxxxxxxxxx\n>")

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

if(website == "3"):
   stopPoint = fileName.index('.')
   prepRev = fileName[0:stopPoint]
   totalName = prepRev + "_rev_cant.xlsx"
   workbook = xlsxwriter.Workbook(totalName)
   worksheet = workbook.add_worksheet()
   worksheet.write(0,0, "Telephone Number")
   worksheet.write(0,1, "Number of Messages Approx.")
   #dont forget to program the zeroes later#
   worksheet.write(0,2, "Number of Scammers")
   worksheet.write(0,3, "Number of Spammers")
   worksheet.write(0,4, "Number of Debt Collectors")
   worksheet.write(0,5, "Number of Hospital")
   worksheet.write(0,6, "Number of People")
   worksheet.write(0,7, "Sentiment")
   
if(website == "EXP1"):
   driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe") 
   stopPoint = fileName.index('.')
   prepRev = fileName[0:stopPoint]
   totalName = prepRev + "_rev_800notes.xlsx"
   workbook = xlsxwriter.Workbook(totalName)
   worksheet = workbook.add_worksheet()
   worksheet.write(0,0, "Telephone Number")
   worksheet.write(0,1, "Number of Messages")
   worksheet.write(0,2, "Number of Pages")
   worksheet.write(0,3, "Number of Scammers")
   worksheet.write(0,4, "Number of Spammers")
   worksheet.write(0,5, "Number of Debt Collectors")
   worksheet.write(0,6, "Number of Hospital")
   worksheet.write(0,7, "Number of People")
   worksheet.write(0,8, "Sentiment")

worksheet.set_column('A:A',13)
col = xl_sheet.col_slice(0,1,10101010)
from xlrd.sheet import ctype_text
#print('(Column #) type:value')
for idx, cell_obj in enumerate(col):
   cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')  
   cell_obj_str = str(cell_obj)
   
   if(numFormat == "1"):
      firstStart = cell_obj_str.index('-')-3
      firstEnd = firstStart + 3   
      secondStart = cell_obj_str.index('-')+1
      secondEnd = secondStart + 3
      thirdStart = cell_obj_str.index('-')+5
      thirdEnd = thirdStart + 4      
      teleWho = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      teleBBB = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      tele800 = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
   
   if(numFormat == "2"):      
      firstStart = cell_obj_str.index('(')+1
      firstEnd = firstStart + 3   
      secondStart = cell_obj_str.index(' ')+1
      secondEnd = secondStart + 3
      thirdStart = cell_obj_str.index('-')+1
      thirdEnd = thirdStart + 4
      teleWho = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      teleBBB = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      tele800 = (cell_obj_str[firstStart:firstEnd] + cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
      
   if(numFormat == "3"):
      teleWho = (cell_obj_str[8:11] + cell_obj_str[11:14] + cell_obj_str[14:18])
      teleBBB = (cell_obj_str[8:11] + cell_obj_str[11:14] + cell_obj_str[14:18])
      tele800 = ("1-" + cell_obj_str[8:11] + "-" + cell_obj_str[11:14] + "-" + cell_obj_str[14:18])
      
   #print('(%s) %s' % (idx, teleWho))
   tnList = teleWho
   worksheet.write(idx+1, 0, "1" + tnList)
   
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
      #print (reqInput)
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
   
   if(website == "3"):
      reqInput = "http://youcantcallus.com/%s" % (tele800)
      urlfile = urllib2.Request(reqInput)
      #print (reqInput)#
      time.sleep(5)
      requestRec = requests.get(reqInput)
      soup = BeautifulSoup(requestRec.content, "lxml")
      noMatch = soup.find(text=re.compile(r"there are 0 comments about this caller"))
      type(noMatch) is str      
      if noMatch is None:
               howMany = soup.find_all('blockquote')
               howManyAreThere = len(howMany)
               worksheet.write(idx+1,1,howManyAreThere)
               scamNum = [ span for span in soup.find_all('span', {'style':'color: #999999; font-size: 12px;'}) if 'scam' in span.text.lower() or 'Scam' in span.text.lower() or 'scams' in span.text.lower() ]
               scamCount = len(scamNum)
               spamNum = [ span for span in soup.find_all('span', {'style':'color: #999999; font-size: 12px;'}) if 'spam' in span.text.lower() or 'Spam' in span.text.lower() or 'spams' in span.text.lower() or 'Survey' in span.text.lower() or 'Telemarketer' in span.text.lower() or 'Political Campaign' in span.text.lower() ]
               spamCount = len(spamNum)     
               debtNum = [ span for span in soup.find_all('span', {'style':'color: #999999; font-size: 12px;'}) if 'debt' in span.text.lower() or 'Debt' in span.text.lower() or 'credit' in span.text.lower() ]
               debtCount = len(debtNum)
               personNum = [ span for span in soup.find_all('span', {'style':'color: #999999; font-size: 12px;'}) if 'Individual' in span.text.lower() or 'Person' in span.text.lower() or 'Human' in span.text.lower() ]
               personCount = len(personNum)               
               hospitalNum = [ span for span in soup.find_all('span', {'style':'font-size:14px; margin:10px; overflow:hidden'}) if 'hospital' in span.text.lower() or 'Hospital' in span.text.lower() ]
               hospitalCount = len(hospitalNum)
               if hospitalCount > 0:
                  hospitalCount+9999
               searchTerms = {'Scam':scamCount,'Spam':spamCount,'Debt Collector':debtCount,'Hospital':hospitalCount,'Person':personCount}
               sentiment = max(searchTerms, key=searchTerms.get)
               worksheet.write(idx+1,2,scamCount)
               worksheet.write(idx+1,3,spamCount)
               worksheet.write(idx+1,4,debtCount)
               worksheet.write(idx+1,5,hospitalCount)
               worksheet.write(idx+1,6,personCount)
               worksheet.write(idx+1,7,sentiment)
               if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0 and personCount == 0:
                  worksheet.write(idx+1,7,"No Entries Detected")      
                  
   if(website == "EXP1"):   
      driver.get('http://800notes.com/Phone.aspx/%s' % (tele800))
      delay = 4
      time.sleep(4)
      requestRec = driver.page_source
      soup = BeautifulSoup(requestRec, "lxml")
      noMatch = soup.find(text=re.compile(r"Report the call using the form"))
      #print (noMatch)
      soup.prettify()
      #print(requestRec.content)
      type(noMatch) is str
      block = soup.find(text=re.compile(r"OctoNet HTTP filter"))
      extrablock = soup.find(text=re.compile(r"returning an unknown error"))
      type(block) is str 
      type(extrablock) is str 
      if(block is not None or extrablock is not None):
         print("\n Damn. Gimme an hour to fix this.")
         time.sleep(5400)
      
      if noMatch is None:
               driver.get('http://800notes.com/Phone.aspx/%s/10000' % (tele800))
               curSite = driver.current_url
               pageExist = soup.find("div", class_="oos_pager")
               type(pageExist) is str 
               if(pageExist is not None):
                  curBegin = curSite.rfind('/') + 1
                  curEnd = curBegin + 4
                  pageNum = curSite[curBegin:curEnd]
               else:
                  pageNum = 1

               
               numMessages = int(pageNum) - 1
               numMessages = numMessages * 20
               convertNum = str(numMessages)
               thumbs = soup.find_all('a',{'class':'oos_i_thumbDown'})
               thumbsLen = len(thumbs)
               thumbPlus = thumbsLen + int(convertNum)
               worksheet.write(idx+1,1,"About " + str(thumbPlus))
                  
               delay = 3
               driver.get('http://800notes.com/Phone.aspx/%s' % (tele800))

               if(pageExist is not None):
                  if(countitup != pageNum):
                     countitup = countitup + 1
                     driver.get('http://800notes.com/Phone.aspx/%s/%s' % (tele800,countitup))
                     delay = 4
                     scamNum = soup.find_all(text=re.compile(r"Scam"))
                     spamNum = soup.find_all(text=re.compile(r"Call type: Telemarketer"))
                     debtNum = soup.find_all(text=re.compile(r"Call type: Debt Collector"))
                     hospitalNum = soup.find_all(text=re.compile(r"Hospital"))
                     scamCount = len(scamNum) + scamCount
                     spamCount = len(spamNum) + spamCount
                     debtCount = len(debtNum) + debtCount
                     hospitalCount = len(hospitalNum) + hospitalCount
                     block = soup.find(text=re.compile(r"OctoNet HTTP filter"))
		     extrablock = soup.find(text=re.compile(r"returning an unknown error"))
                     type(block) is str 
		     type(extrablock) is str 
                     if(block is not None or extrablock is not None):
                        print("\n Damn. Gimme an hour to fix this.")
                        time.sleep(2000)
                        
                  if hospitalCount > 0:
                     hospitalCount+9999
                  searchTerms = {'Scam':scamCount,'Spam':spamCount,'Debt Collector':debtCount,'Hospital':hospitalCount}
                  sentiment = max(searchTerms, key=searchTerms.get) 
                  worksheet.write(idx+1,3,scamCount)
                  worksheet.write(idx+1,4,spamCount)
                  worksheet.write(idx+1,5,debtCount)
                  worksheet.write(idx+1,6,hospitalCount)
                  worksheet.write(idx+1,7,sentiment)
                  if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0 and personCount == 0:
                     worksheet.write(idx+1,7,"No Entries Detected")
                  countitup = 0
                  scamCount = 0
                  spamCount = 0
                  debtCount = 0
                  hospitalCount = 0
               
               worksheet.write(idx+1,2,pageNum)                  

driver.quit()
workbook.close()

if delMe == 1:
   os.remove(deleteFile)
   print("Temp File Cleaned!\n")

done = True
ding = "\nDing! Job Done!"
uni = unicode( ding, "utf-8")
print (ding)