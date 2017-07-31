
powershell -Command "(New-Object Net.WebClient).DownloadFile('https://download.microsoft.com/download/7/9/6/796EF2E4-801B-4FC4-AB28-B59FBF6D907B/VCForPython27.msi', 'VCForPython27.msi')"
VCForPython27.msi
easy_install pip
pip install -q -r requirements.txt
