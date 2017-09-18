[![forthebadge](http://forthebadge.com/images/badges/compatibility-betamax.svg)](http://forthebadge.com)

# Haircuttery

A program I wrote for as a personal project, meant to scrape Better Business Bureau and 800notes with an excel file.

## Sites Supported

http://whoscall.in/

https://www.bbb.org/

http://800notes.com/

https://www.shouldianswer.com/

https://people.yellowpages.com

https://www.yellowpages.com

# How to Use

## Windows Users:

Just drag the file you would like to work on into "DropOnMe.bat." You'll be asked to choose a site to scrape, or type "A" to do an all in one scrape and bundle operation with the important info seperated.

## Linux/Mac Users

Make sure you have the Dependencies listed below, and run HairCut.py via either the terminal or double clicking, should your distro support that.

### FileSplit

Dropping Files into DropOnMe.bat will seperate them into 1000 cell files which can be used to prevent lockout on certain sites.

### Chrome Support

1. Download the appropriate executable for your operating system <https://sites.google.com/a/chromium.org/chromedriver/downloads>

2. Copy the File to one of three places:

  - C:\ aka the root of your Hard Drive
  - The Working Directory of Haircuttery
  - Or a directory of your choice by creating a file called "chrome.ini" with the directory and executable on one line.

## Excel or CSV format

Check Example.xlsx for an example of a correctly formatted input file.

## Dependencies:

### Windows Users

Run requirements.bat to install necessary files and the Python Compiler.

### General

- Chrome with Chrome Webdriver.

### Python

- xlrd
- xlwt
- bs4
- requests
- flask
- numpy
- regex
- pyopenssl
- ndg-httpsclient
- pyasn1
- xlsxwriter
- lxml
- selenium

pip can fetch these with the command "pip install -r requirements.txt"

# This program is no longer being worked on by me, thanks for all the fish.

(ಥ﹏ಥ)

I'll accept pull requests for new features though.

## Credits:

StackOverflow, RTFM is too hard for me so you guys rock.

FGAO22, for an awesome example and basis. <https://github.com/fgao22/GSD>

rprasad, for his excellent XLRD guide. <https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/>

Diego Rosales, Helpful Programming Friend. My Rubber Duck and Siamese mind.
