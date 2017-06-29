from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from Harvard import Excel2CSV, enumColumn, TimeOutHandler
from IPython.display import HTML
from os.path import join, dirname, abspath
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from shutil import copyfile
from shutil import move
from xlrd.sheet import ctype_text
import csv
import glob
import itertools
import numpy as np
import os
import re
import requests
import selenium.webdriver.support.ui as ui
import sys
import threading
import time
import urllib2
import xlrd
import xlsxwriter
import xlwt

def loading():
    for s in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        if(breaker == 1):
            break
        sys.stdout.write('\rloading ' + s)
        sys.stdout.flush()
        time.sleep(0.1)

breaker = 0
countitup = 1
debtCount = 0
delMe = 0
done = False
hospitalCount = 0
scamCount = 0
spamCount = 0
w = 1
x = 2
y = 3
z = 4

book = xlwt.Workbook(encoding="utf-8")
headers = {
    'User-Agent': 'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}
worksheet = book.add_sheet("Results", cell_overwrite_ok=True)
file_paths = sys.argv[1:2]
draganddrop = ''.join(file_paths)
webType = sys.argv[2:3]
dragNDrop2 = ''.join(webType)
numType = sys.argv[3:4]
dragNDrop3 = ''.join(numType)

if draganddrop == "":
    fileName = raw_input("\nInput the file with extension\n>")
else:
    fileOnly = draganddrop.rfind('\\') + 1
    fileName = draganddrop[fileOnly:]

stopPoint = fileName.index('.')
prepRev = fileName[stopPoint:]
preName = fileName[:stopPoint]

reload(sys)
sys.setdefaultencoding('utf')

if prepRev == ".csv":
    totalName = preName + '.xlsx'
    excelFile = xlsxwriter.Workbook(totalName)
    excelFile.close()
    excelFile = xlsxwriter.Workbook(totalName)
    worksheet = excelFile.add_worksheet()
    enumColumn(fileName=fileName,worksheet=worksheet)
    excelFile.close()
    fileName = (preName + '.xlsx')
    delMe = 1
    print("Temporary Convert to xlsx done.\n")

fname = join(dirname(abspath('__file__')), '%s' % fileName)

xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

if dragNDrop2 == "":
    website = raw_input(
        "Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes\n>")
else:
    website = dragNDrop2

if dragNDrop3 == "":
    numFormat = raw_input(
        "Which format?\n1 for xxx-xxx-xxxx, 2 for (xxx) xxx-xxxx, 3 for xxxxxxxxxx\n>")
else:
    numFormat = dragNDrop3

g = threading.Thread(target=loading)
g.start()

if(website == "1"):
    stopPoint = fileName.index('.')
    prepRev = fileName[0:stopPoint]
    totalName = prepRev + "_rev_who.xlsx"
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Telephone Number")
    worksheet.write(0, 1, "# of Messages")
    worksheet.write(0, 2, "Does it Appear?")
    worksheet.write(0, 3, "Number of Scammers")
    worksheet.write(0, 4, "Number of Spammers")
    worksheet.write(0, 5, "Number of Debt Collectors")
    worksheet.write(0, 6, "Number of Hospital")
    worksheet.write(0, 7, "Sentiment")
    siteType = "_rev_who.xlsx"

if(website == "2"):
    if(os.path.exists(r"C:/chromedriver.exe") or os.path.exists(r"chromedriver.exe")):
        driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe")
        driver.set_page_load_timeout(5006900)
        stopPoint = fileName.index('.')
        prepRev = fileName[0:stopPoint]
        totalName = prepRev + "_rev_BBB.xlsx"
        workbook = xlsxwriter.Workbook(totalName)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Telephone Number")
        worksheet.write(0, 1, "Acreditted")
        siteType = "_rev_BBB.xlsx"
    else:
        breaker = 1
        print("\nPlease refer to the Readme, you don't have chromedriver.exe in 'C:\chromedriver'")
        time.sleep(15)
        sys.exit()

if(website == "3"):
    if(os.path.exists(r"C:/chromedriver.exe") or os.path.exists(r"chromedriver.exe")):
        driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe")
        driver.set_page_load_timeout(600)
        stopPoint = fileName.index('.')
        prepRev = fileName[0:stopPoint]
        totalName = prepRev + "_rev_800notes.xlsx"
        workbook = xlsxwriter.Workbook(totalName)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Telephone Number")
        worksheet.write(0, 1, "Approximate Number of Messages")
        worksheet.write(0, 2, "Number of Pages")
        worksheet.write(0, 3, "Number of Scammers")
        worksheet.write(0, 4, "Number of Spammers")
        worksheet.write(0, 5, "Number of Debt Collectors")
        worksheet.write(0, 6, "Number of Hospital")
        worksheet.write(0, 7, "Sentiment")
        siteType = "_rev_800notes.xlsx"
    else:
        breaker = 1
        print("\nPlease refer to the Readme, you don't have chromedriver.exe in 'C:\chromedriver'")
        time.sleep(15)
        sys.exit()

worksheet.set_column('A:A', 13)
col = xl_sheet.col_slice(0, 1, 10101010)
for idx, cell_obj in enumerate(col):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    cell_obj_str = str(cell_obj)

    if(numFormat == "1"):
        firstStart = cell_obj_str.index('-') - 3
        firstEnd = firstStart + 3
        secondStart = cell_obj_str.index('-') + 1
        secondEnd = secondStart + 3
        thirdStart = cell_obj_str.index('-') + 5
        thirdEnd = thirdStart + 4
        teleWho = (cell_obj_str[firstStart:firstEnd] +
                   cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
        teleBBB = (cell_obj_str[firstStart:firstEnd] +
                   cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
        tele800 = ("1-" + cell_obj_str[firstStart:firstEnd] + "-" +
                   cell_obj_str[secondStart:secondEnd] + "-" + cell_obj_str[thirdStart:thirdEnd])

    if(numFormat == "2"):
        firstStart = cell_obj_str.index('(') + 1
        firstEnd = firstStart + 3
        secondStart = cell_obj_str.index(' ') + 1
        secondEnd = secondStart + 3
        thirdStart = cell_obj_str.index('-') + 1
        thirdEnd = thirdStart + 4
        teleWho = (cell_obj_str[firstStart:firstEnd] +
                   cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
        teleBBB = (cell_obj_str[firstStart:firstEnd] +
                   cell_obj_str[secondStart:secondEnd] + cell_obj_str[thirdStart:thirdEnd])
        tele800 = ("1-" + cell_obj_str[firstStart:firstEnd] + "-" +
                   cell_obj_str[secondStart:secondEnd] + "-" + cell_obj_str[thirdStart:thirdEnd])

    if(numFormat == "3"):
        teleWho = (cell_obj_str[8:11] +
                   cell_obj_str[11:14] + cell_obj_str[14:18])
        teleBBB = (cell_obj_str[8:11] +
                   cell_obj_str[11:14] + cell_obj_str[14:18])
        tele800 = ("1-" + cell_obj_str[8:11] + "-" +
                   cell_obj_str[11:14] + "-" + cell_obj_str[14:18])

    worksheet.write(idx + 1, 0, "1" + teleWho)

    if(website == "1"):
        reqInput = "http://whoscall.in/1/%s/" % (teleWho)
        urlfile = urllib2.Request(reqInput)
        time.sleep(1)
        requestRec = requests.get(reqInput)
        soup = BeautifulSoup(requestRec.content, "lxml")
        noMatch = soup.find(text=re.compile(
            r"no reports yet on the phone number"))
        type(noMatch) is str
        if noMatch is None:
            worksheet.write(idx + 1, 2, "Got a hit")
            howMany = soup.find_all('img', {'src': '/default-avatar.gif'})
            howManyAreThere = len(howMany)
            worksheet.write(idx + 1, 1, howManyAreThere)
            scamNum = [div for div in soup.find_all('div', {'style': 'font-size:14px; margin:10px; overflow:hidden'})
                       if 'scam' in div.text.lower() or 'Scam' in div.text.lower() or 'scams' in div.text.lower()]
            scamCount = len(scamNum)
            spamNum = [div for div in soup.find_all('div', {'style': 'font-size:14px; margin:10px; overflow:hidden'})
                       if 'spam' in div.text.lower() or 'Spam' in div.text.lower() or 'spams' in div.text.lower()]
            spamCount = len(spamNum)
            debtNum = [div for div in soup.find_all('div', {'style': 'font-size:14px; margin:10px; overflow:hidden'})
                       if 'debt' in div.text.lower() or 'Debt' in div.text.lower() or 'credit' in div.text.lower()]
            debtCount = len(debtNum)
            hospitalNum = [div for div in soup.find_all(
                'div', {'style': 'font-size:14px; margin:10px; overflow:hidden'}) if 'hospital' in div.text.lower() or 'Hospital' in div.text.lower()]
            hospitalCount = len(hospitalNum)
            if hospitalCount > 0:
                hospitalCount + 9999
            searchTerms = {'Scam': scamCount, 'Spam': spamCount,
                           'Debt Collector': debtCount, 'Hospital': hospitalCount}
            sentiment = max(searchTerms, key=searchTerms.get)
            worksheet.write(idx + 1, 3, scamCount)
            worksheet.write(idx + 1, 4, spamCount)
            worksheet.write(idx + 1, 5, debtCount)
            worksheet.write(idx + 1, 6, hospitalCount)
            worksheet.write(idx + 1, 7, sentiment)
            if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0:
                worksheet.write(idx + 1, 7, "No Entries Detected")

    if(website == "2"):
        driver.get('https://www.bbb.org/en/us/search?inputText=' +
                   teleBBB + '&locationText=&locationLatLng=&page=1')
        time.sleep(1)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")
        Hit = soup.find_all('aside', {'class': 'search-result__aside'})
        driver.get('https://www.bbb.org/en/us/search?accreditedFilter=1&inputText=' +
                   teleBBB + '&locationText=&locationLatLng=&page=1')
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")
        Badge = soup.find_all('aside', {'class': 'search-result__aside'})
        if len(Hit) != 0:
            worksheet.write(idx + 1, 1, "Got a Hit")
        if len(Badge) != 0:
            worksheet.write(idx + 1, 1, "Is Accredited")

    if(website == "3"):
        try:
            driver.get('http://800notes.com/Phone.aspx/%s' % (tele800))
        except TimeoutException as ex:
            TimeOutHandler(driver=driver,worksheet=worksheet,webdriver=webdriver)
            break
        delay = 4
        time.sleep(4)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")
        noMatch = soup.find(text=re.compile(r"Report the call using the form"))
        soup.prettify()
        type(noMatch) is str
        block = soup.find(text=re.compile(r"OctoNet HTTP filter"))
        extrablock = soup.find(text=re.compile(r"returning an unknown error"))
        extrablock = soup.find(text=re.compile(r"Gateway time-out"))
        type(block) is str
        type(extrablock) is str
        if(block is not None or extrablock is not None):
            print("\n Damn. Gimme an hour to fix this.")
            time.sleep(5000)

        if noMatch is None:
            try:
                driver.get('http://800notes.com/Phone.aspx/%s/10000' %
                           (tele800))
            except TimeoutException as ex:
                TimeOutHandler(driver=driver,worksheet=worksheet,webdriver=webdriver)
                break
            curSite = driver.current_url
            pageExist = soup.find("a", class_="oos_i_thumbDown")
            type(pageExist) is str
            if(pageExist is not None):
                curBegin = curSite.rfind('/') + 1
                curEnd = curBegin + 4
                pageNum = curSite[curBegin:curEnd]
            else:
                pageNum = 1
            if(curSite.count("/") < 5):
                pageNum = 1
            numMessages = int(pageNum) - 1
            numMessages = numMessages * 20
            convertNum = str(numMessages)
            thumbs = soup.find_all('a', {'class': 'oos_i_thumbDown'})
            thumbsLen = len(thumbs)
            thumbPlus = thumbsLen + int(convertNum)
            worksheet.write(idx + 1, 1, thumbPlus)
            delay = 3
            if(pageExist is not None):
                while(int(countitup) != int(pageNum) + 1):
                    try:
                        if(countitup == 1):
                            driver.get(
                                'http://800notes.com/Phone.aspx/{}'.format(tele800))
                        else:
                            driver.get(
                                'http://800notes.com/Phone.aspx/{}/{}/'.format(tele800, countitup))
                    except TimeoutException as ex:
                        TimeOutHandler(driver=driver,worksheet=worksheet,webdriver=webdriver)
                        break
                    delay = 4
                    requestRec = driver.page_source
                    soup = BeautifulSoup(requestRec, "lxml")
                    countitup = int(countitup) + 1
                    if (countitup % 2 == 0):
                        time.sleep(5)
                    else:
                        time.sleep(4)
                    scamNum = soup.find_all(
                        'div', class_="oos_contletBody", text=re.compile(r"Scam", flags=re.IGNORECASE))
                    spamNum = soup.find_all(
                        text=re.compile(r"Call type: Telemarketer"))
                    debtNum = soup.find_all(
                        text=re.compile(r"Call type: Debt collector"))
                    hospitalNum = soup.find_all('div', class_="oos_contletBody", text=re.compile(
                        r"Hospital", flags=re.IGNORECASE))
                    scamCount = len(scamNum) + scamCount
                    spamCount = len(spamNum) + spamCount
                    debtCount = len(debtNum) + debtCount
                    hospitalCount = len(hospitalNum) + hospitalCount
                    block = soup.find(text=re.compile(r"OctoNet HTTP filter"))
                    extrablock = soup.find(text=re.compile(
                        r"returning an unknown error"))
                    extrablock = soup.find(
                        text=re.compile(r"Gateway time-out"))
                    type(block) is str
                    type(extrablock) is str
                    if(block is not None or extrablock is not None):
                        print("\n Damn. Gimme an hour to fix this.")
                        time.sleep(5000)
                if hospitalCount > 0:
                    hospitalCount + 9999
                searchTerms = {'Scam': scamCount, 'Spam': spamCount,
                               'Debt Collector': debtCount, 'Hospital': hospitalCount}
                sentiment = max(searchTerms, key=searchTerms.get)
                worksheet.write(idx + 1, 3, scamCount)
                worksheet.write(idx + 1, 4, spamCount)
                worksheet.write(idx + 1, 5, debtCount)
                worksheet.write(idx + 1, 6, hospitalCount)
                worksheet.write(idx + 1, 7, sentiment)
                if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0:
                    worksheet.write(idx + 1, 7, "No Entries Detected")
            countitup = 1
            debtCount = 0
            hospitalCount = 0
            scamCount = 0
            spamCount = 0
            worksheet.write(idx + 1, 2, int(pageNum))

if(website == "2" or website == "3"):
    driver.close()
workbook.close()
prepRev = preName + '_temp.csv'
Excel2CSV(totalName, "Sheet1", prepRev)

if not os.path.exists(preName):
    os.makedirs("WorkingDir/" + preName)
if prepRev == ".csv":
    totalName = preName + prepRev
else:
    totalName = preName + ".xlsx"
copyfile(totalName, "WorkingDir/" + preName + '/' + totalName)
move(preName + siteType, "WorkingDir/" + preName + '/' + preName + siteType)
move(preName + "_temp.csv", "WorkingDir/" +
     preName + '/' + preName + "_temp.csv")

done = True
print ("\nDing! Job Done!")
