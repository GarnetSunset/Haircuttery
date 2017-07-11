[![forthebadge](http://forthebadge.com/images/badges/compatibility-betamax.svg)](http://forthebadge.com)

# Haircuttery

A program I wrote for as a personal project, meant to scrape Better Business Bureau and 800notes with an excel file.

## How to Use

### Windows Users:

Just drag the file you would like to work on into "DropOnMe.bat." You'll be asked to choose a site to scrape.

### Linux/Mac Users

Make sure you have the Dependencies listed below, and run HairCut.py via either the terminal or double clicking, should your distro support that.

### FileSplit

Dropping Files into DropOnMe.bat will seperate them into 1000 cell files which can be used to prevent lockout on certain sites.

### Windows Users

Download and Install this: <https://www.microsoft.com/en-us/download/details.aspx?id=44266>

#### Chrome Support

1. Download the appropriate executable for your operating system <https://sites.google.com/a/chromium.org/chromedriver/downloads>

2. Copy the File to one of three places:

  - C:\ aka the root of your Hard Drive
  - The Working Directory of Haircuttery
  - Or a directory of your choice by creating a file called "chrome.ini" with the directory and executable on one line.

## Dependencies:

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

## Credits:

FGAO22, for an awesome example and basis. <https://github.com/fgao22/GSD>

rprasad, for his XLRD guide. <https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/>

Diego Rosales, Helpful Programming Friend. My Rubber Duck and Siamese mind.
