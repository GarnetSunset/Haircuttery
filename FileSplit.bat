@ECHO OFF

copy /b NUL tempFile

cls

python "FileSplit.py" %*
pause
cls

set /p folderName=<tempFile

del /f tempFile

cd WorkingDir

cd %folderName%

Set COUNTER=1

set fileName="%folderName%_%COUNTER%.xlsx"

SET /P webSite=Input 1 for whoscall.in results, input 2 for BBB, input 3 for 800Notes:

SET /P tnFormat=Which format? 1 for xxx-xxx-xxxx, 2 for (xxx) xxx-xxxx, 3 for xxxxxxxxxx:

:HairCut

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

pause