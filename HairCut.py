from __future__ import print_function
from bs4 import BeautifulSoup
from collections import defaultdict
from Harvard import Excel2CSV, enumColumn
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
import logging
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

# Wait out my mistake.


def blocked():
    block = soup.find(text=re.compile(r"OctoNet HTTP filter"))
    extrablock = soup.find(text=re.compile(
        r"returning an unknown error"))
    extrablock = soup.find(
        text=re.compile(r"Gateway time-out"))
    type(block) is str
    type(extrablock) is str
    if(block is not None or extrablock is not None):
        print("\n Ugh. I'm gonna go talk to the host of the site real quick. Should take an hour or two.")
        time.sleep(7200)

# Break if no chromedriver.


def breaker():
    done = True
    print("\nPlease refer to the Readme, you don't have chromedriver.exe in 'C:\chromedriver'")
    time.sleep(15)
    sys.exit()

# Loading Animation that plays when the user is running a file.


def loading():
    for s in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + s)
        sys.stdout.flush()
        time.sleep(0.1)

# TimeoutHandler that takes care of webDriver fails.


def TimeOutHandler(driver, webdriver, worksheet):
    driver.close()
    driver = webdriver.Chrome()
    worksheet.write(idx + 1, 7, "Timeout Exception")


bbbEnd = "&locationText=&locationLatLng=&page=1"
bbbUrl = "https://www.bbb.org/en/us/search?inputText="
bbbUrlAC = "https://www.bbb.org/en/us/search?accreditedFilter=1&inputText="
breaker = 0
countitup = 1
debtCount = 0
delMe = 0
done = False
hospitalCount = 0
numFormat = "3"
scamCount = 0
spamCount = 0

# Create a UTF-8 Workbook.
book = xlwt.Workbook(encoding="utf-8")

# Assign a User-Agent to python.
headers = {
    'User-Agent': 'Chrome/39.0.2171.95 Safari/537.36 AppleWebKit/537.36 (KHTML, like Gecko)'}

# Create a worksheet named "Results".
worksheet = book.add_sheet("Results", cell_overwrite_ok=True)

# Join the dragged files to create strings.
dragNDrop = ''.join(sys.argv[1:2])
dragNDrop2 = ''.join(sys.argv[2:3])

# Was a file dragged onto the Batch file?
# If not the string "dragNDrop" will be empty and the user will be prompted.
if dragNDrop == "":
    fileName = raw_input("\nInput the file with extension\n>")
else:
    # Obtain the fileName only by removing the directory name.
    fileOnly = dragNDrop.rfind('\\') + 1
    fileName = dragNDrop[fileOnly:]

# Was a site given in the Batch file?
# If not the string "dragNDrop2" will be empty and the user will be prompted.
if dragNDrop2 == "":
    website = raw_input(
        "Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes\n>")
else:
    website = dragNDrop2

# Find the period in the file, which determines the prepRev or extension, and the fileName.
stopPoint = fileName.index('.')
prepRev = fileName[stopPoint:]
preName = fileName[:stopPoint]

# Make sure we're still encoding in UTF. Don't want any mistakes now, do we?
reload(sys)
sys.setdefaultencoding('utf')

# Is the extension CSV? If so we'll convert it to xlsx.
if prepRev == ".csv":
    totalName = preName + '.xlsx'
    excelFile = xlsxwriter.Workbook(totalName)
    excelFile.close()
    # You must close the file because opening fails otherwise.
    excelFile = xlsxwriter.Workbook(totalName)
    worksheet = excelFile.add_worksheet()
    enumColumn(fileName=fileName, worksheet=worksheet)
    excelFile.close()
    fileName = (preName + '.xlsx')
    print("Temporary Convert to xlsx done.\n")

# Get ready for XLRD to parse the original file (or the converted one).
fname = join(dirname(abspath('__file__')), '%s' % fileName)

# Parse it XLRD!
xl_workbook = xlrd.open_workbook(fname)
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

# If the user types "d" for the website choice, they will be prompted again, but, this time given debug info.
if(website == "d"):
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')
    website = raw_input(
        "Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes\n>")
    logging.basicConfig(level=logging.DEBUG)
    logging.debug('Only shown in debug mode')

# Start the little spinny animation.
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
    if(os.path.exists(r"C:/chromedriver.exe") or os.path.isfile('chromedriver.exe')):
        if(os.path.exists(r"C:/chromedriver.exe")):
            driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe")
        if(os.path.isfile('chromedriver.exe')):
            driver = webdriver.Chrome(executable_path='chromedriver.exe')
        driver.set_page_load_timeout(600)
        stopPoint = fileName.index('.')
        prepRev = fileName[0:stopPoint]
        totalName = prepRev + "_rev_BBB.xlsx"
        workbook = xlsxwriter.Workbook(totalName)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Telephone Number")
        worksheet.write(0, 1, "Acreditted")
        siteType = "_rev_BBB.xlsx"
    else:
        breaker()

if(website == "3"):
    if(os.path.exists(r"C:/chromedriver.exe") or os.path.isfile('chromedriver.exe')):
        if(os.path.exists(r"C:/chromedriver.exe")):
            driver = webdriver.Chrome(executable_path=r"C:/chromedriver.exe")
        if(os.path.isfile('chromedriver.exe')):
            driver = webdriver.Chrome(executable_path='chromedriver.exe')
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
        breaker()

# Set column to A:A, the first column.
worksheet.set_column('A:A', 13)
# Read the slice from the first cell to the last accessible row in Excel.
col = xl_sheet.col_slice(0, 1, 1048576)
# Read each string line by line.
for idx, cell_obj in enumerate(col):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    cell_obj_str = str(cell_obj)

    # Does a dash, parenthesis, or none of those exist? That will decide the numFormat.
    if "-" in cell_obj_str:
        numFormat = "1"

    if "(" in cell_obj_str:
        numFormat = "2"

    # Cut the numbers to their appropriate format.

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

    # WhosCallinScrapes using the python Requests library. Nice and clean.
    if(website == "1"):
        reqInput = "http://whoscall.in/1/%s/" % (teleWho)
        time.sleep(1)
        urlfile = urllib2.Request(reqInput)
        requestRec = requests.get(reqInput)
        soup = BeautifulSoup(requestRec.content, "lxml")
        noMatch = soup.find(text=re.compile(
            r"no reports yet on the phone number"))
        type(noMatch) is str
        if noMatch is None:
            worksheet.write(idx + 1, 2, "Got a hit")

            # Check for number of comments.
            howMany = soup.find_all('img', {'src': '/default-avatar.gif'})
            howManyAreThere = len(howMany)
            worksheet.write(idx + 1, 1, howManyAreThere)

            # Search for text on the sites that indicates their sentiment and generate the top response.
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
            worksheet.write(idx + 1, 3, scamCount)
            worksheet.write(idx + 1, 4, spamCount)
            worksheet.write(idx + 1, 5, debtCount)
            worksheet.write(idx + 1, 6, hospitalCount)

            # Hospitals are important to look at, so I boost them.
            if hospitalCount > 0:
                hospitalCount + 9999

            searchTerms = {'Scam': scamCount, 'Spam': spamCount,
                           'Debt Collector': debtCount, 'Hospital': hospitalCount}
            sentiment = max(searchTerms, key=searchTerms.get)
            worksheet.write(idx + 1, 7, sentiment)

            # Stalemate?
            if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0:
                worksheet.write(idx + 1, 7, "No Entries Detected")

    if(website == "2"):

        # Selenium, get that site for me! (bbbUrl + bbbEnd are defined above)
        driver.get(bbbUrl +
                   teleBBB + bbbEnd)
        time.sleep(1)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")
        Hit = soup.find_all('aside', {'class': 'search-result__aside'})

        # Cloned the previous section, but, with changes to the URL via bbbUrlAC.
        driver.get(bbbUrlAC +
                   teleBBB + bbbEnd)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")
        Badge = soup.find_all('aside', {'class': 'search-result__aside'})
        if len(Hit) != 0:
            worksheet.write(idx + 1, 1, "Got a Hit")
        if len(Badge) != 0:
            worksheet.write(idx + 1, 1, "Is Accredited")

    if(website == "3"):
        # 800Notes, the big one.
        try:
            driver.get('http://800notes.com/Phone.aspx/%s' % (tele800))
        except TimeoutException as ex:
            TimeOutHandler(driver=driver, worksheet=worksheet,
                           webdriver=webdriver)
            break
        time.sleep(2)
        requestRec = driver.page_source
        soup = BeautifulSoup(requestRec, "lxml")

        # This entry doesn't exist if this regex succeeds.
        noMatch = soup.find(text=re.compile(r"Report the call using the form"))
        soup.prettify()
        type(noMatch) is str

        # Make sure we don't get blocked, and if we do, wait it out.
        blocked()

        if noMatch is None:
            try:
                driver.get('http://800notes.com/Phone.aspx/%s/10000' %
                           (tele800))
            except TimeoutException as ex:
                TimeOutHandler(driver=driver, worksheet=worksheet,
                               webdriver=webdriver)
                break
            blocked()
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
            time.sleep(2)
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
                        TimeOutHandler(
                            driver=driver, worksheet=worksheet, webdriver=webdriver)
                        break
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
                    blocked()
                worksheet.write(idx + 1, 3, scamCount)
                worksheet.write(idx + 1, 4, spamCount)
                worksheet.write(idx + 1, 5, debtCount)
                worksheet.write(idx + 1, 6, hospitalCount)

                if hospitalCount > 0:
                    hospitalCount + 9999

                searchTerms = {'Scam': scamCount, 'Spam': spamCount,
                               'Debt Collector': debtCount, 'Hospital': hospitalCount}
                sentiment = max(searchTerms, key=searchTerms.get)
                worksheet.write(idx + 1, 7, sentiment)
                if scamCount == 0 and spamCount == 0 and debtCount == 0 and hospitalCount == 0:
                    worksheet.write(idx + 1, 7, "No Entries Detected")
            countitup = 1
            debtCount = 0
            hospitalCount = 0
            scamCount = 0
            spamCount = 0
            worksheet.write(idx + 1, 2, int(pageNum))

# Close up Shop!
workbook.close()
prepRev = preName + '_temp.csv'
Excel2CSV(totalName, "Sheet1", prepRev)

# Determine if file was dragged or not for creation of Dirs.
if dragNDrop == "":
    if not os.path.exists("WorkingDir"):
        os.makedirs("WorkingDir")
    if not os.path.exists("WorkingDir/" + preName):
        os.makedirs("WorkingDir/" + preName)

# Was the file originially a CSV?
if prepRev == ".csv":
    totalName = preName + prepRev
else:
    totalName = preName + ".xlsx"

# If we haven't already moved all of the files, here we go.
if dragNDrop == "":
    copyfile(totalName, "WorkingDir/" + preName + '/' + totalName)
    move(preName + siteType, "WorkingDir/" +
         preName + '/' + preName + siteType)
    move(preName + "_temp.csv", "WorkingDir/" +
         preName + '/' + preName + "_temp.csv")

# End Animation.
done = True
print ("\nDing! Job Done!")
