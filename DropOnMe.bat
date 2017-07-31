@ECHO OFF

cls

python "FileSplit.py" %*

cls

if exist "tempSmall.log" (goto JustGoForIt)

set /p folderName=<"tempFile.log"

if exist "tempFile.log" (del /f "tempFile.log")

cd /d WorkingDir

cd /d %folderName%

Set COUNTER=1

cls

echo Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes,
SET /P webSite=input 4 for ShouldIAnswer, input 5 for YellowPages:

:HairCut

set fileName="%folderName%_%COUNTER%.xlsx"

if exist %fileName% (

python "HairCut.py" %fileName% %webSite% %tnFormat%

) else (

ECHO That's All!

goto youWin

)

set /A COUNTER=COUNTER+1

goto Haircut

:youWin

cd ..

EXIT /B

:JustGoForIt

set /p folderName=<"tempSmall.log"

del /f "tempSmall.log"

cd /d WorkingDir

cd /d %folderName%

cls

python "Haircut.py" %*

EXIT /B
