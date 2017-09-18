powershell -Command "(New-Object Net.WebClient).DownloadFile('https://download.microsoft.com/download/7/9/6/796EF2E4-801B-4FC4-AB28-B59FBF6D907B/VCForPython27.msi', 'VCForPython27.msi')"
VCForPython27.msi
del VCForPython27.msi
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://chromedriver.storage.googleapis.com/2.32/chromedriver_win32.zip', 'chromedriver_win32.zip')"
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://stahlworks.com/dev/unzip.exe', 'unzip.exe')"
unzip chromedriver_win32.zip
del chromedriver_win32.zip
del unzip.exe
easy_install pip
pip install -q -r requirements.txt
