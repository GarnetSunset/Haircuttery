@ECHO OFF

setlocal EnableDelayedExpansion

cls

python "FileSplit.py" %*

cls

if exist "tempFile.log" (

set /p folderName=<"tempFile.log"

del /f "tempFile.log")

if exist "tempSmall.log" (

set /p folderName=<"tempSmall.log"

del /f "tempSmall.log")

cd /d WorkingDir

cd /d %folderName%

Set COUNTER=1

cls

echo Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes,
SET /P webSite=input 4 for ShouldIAnswer, input 5 for YellowPages, or A for All:

:HairCut

set fileName="%folderName%.xlsx"

if exist %fileName% (

if "%webSite%" == "A" (

python "Haircut.py" %fileName% 1

python "Haircut.py" %fileName% 2

python "Haircut.py" %fileName% 3

python "Haircut.py" %fileName% 4

python "Haircut.py" %fileName% 5

python "Haircut.py" %fileName% A

) else (

python "HairCut.py" %fileName% %webSite%

)

) else (

ECHO That's All!

goto youWin

)

:youWin

cd ..

EXIT /B
