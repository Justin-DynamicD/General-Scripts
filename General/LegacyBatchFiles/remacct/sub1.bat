REM %1 = Account Name
REM %2 = Share Volume
REM %3 = Backup Location

:MVDATA
set srv=
set drv=
set pth=
IF EXIST tmp.txt erase tmp.txt
IF EXIST tmp2.txt erase tmp2.txt
IF '%2'=='' GOTO CLNUP
IF '%3'=='' GOTO CLNUP
IF NOT EXIST %2 GOTO CLNUP
echo moving %2

md %3\%1
IF ERRORLEVEL 1 GOTO CLNUP
scopy %2 %3\%1 /s


rmtshare %2 | FIND "Path" > tmp.txt
FOR /F "tokens=2" %%A IN (tmp.txt) DO Echo %%A > tmp2.txt
echo %2 > tmp.txt
FOR /F "tokens=1 delims=\" %%A in (tmp.txt) DO SET srv=%%A
FOR /F "tokens=1 delims=:" %%A in (tmp2.txt) DO SET drv=%%A
FOR /F "tokens=2 delims=:" %%A in (tmp2.txt) DO SET pth=%%A

rmtshare %2 /delete
net use W: \\%srv%\%drv%$
deltree /y W:%pth%
net use W: /delete /y

:CLNUP
set srv=
set drv=
set pth=
rem IF EXIST tmp.txt erase tmp.txt
rem IF EXIST tmp2.txt erase tmp2.txt

:END
