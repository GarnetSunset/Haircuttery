pip install -q -r requirements.txt
pip install python-dateutil --upgrade
cd downloader
chromeDriver.py
del chromeDriver.py
move chromedriver.exe ..
cd ..
rmdir /Q/S downloader