@Echo Off
rem %1 is the username
rem %2 is the home share or server
rem %3 is the optional original share

:PREPENV
set srv=
set shr=
set ltr=
set pth=
IF EXIST tmp.txt erase tmp.txt
IF EXIST tmp2.txt erase tmp2.txt

:INPUTCHK
IF '%1'=='?' start wordpad readme.wri
IF '%1'=='?' GOTO END
IF not '%OS%'=='Windows_NT' echo This batch file must be run under an NT based system.
IF not '%OS%'=='Windows_NT' goto END
IF '%2'=='' echo Usage: fixshr [username] ([home share] or [Server]) (Optional:[original share])
IF '%2'=='' echo        Run "fixshr ?" for more help
IF '%2'=='' GOTO END
IF exist \\%2\c$ GOTO ADDCHECK
IF not exist %2 echo %2 is neither a valid home share or Server
IF not exist %2 GOTO END

:FINDPATH
rmtshare %2 | FIND "Path" > tmp.txt
if Errorlevel 1 goto REHOME
FOR /F "tokens=2" %%A IN (tmp.txt) DO Echo %%A > tmp2.txt
echo %2 > tmp.txt
FOR /F "tokens=1 delims=\" %%A in (tmp.txt) DO SET srv=%%A
FOR /F "tokens=1 delims=:" %%A in (tmp2.txt) DO SET ltr=%%A
FOR /F "tokens=2 delims=:" %%A in (tmp2.txt) DO SET pth=%%A
goto MAKECHNGE

:REHOME
echo %2 > tmp.txt
FOR /F "tokens=1 delims=\" %%A in (tmp.txt) DO SET srv=%%A
FOR /F "tokens=2 delims=\" %%A in (tmp.txt) DO SET shr=%%A
FOR /F "tokens=1* delims=%srv%" %%A in (tmp.txt) DO SET pth=%%B
rmtshare \\%srv%\%shr% | FIND "Path" > tmp.txt
FOR /F "tokens=2" %%A IN (tmp.txt) DO Echo %%A > tmp2.txt
FOR /F "tokens=1 delims=:" %%A in (tmp2.txt) DO SET ltr=%%A
goto MAKECHNGE

:MAKECHNGE
echo Cleaning \\%srv%\%ltr%$%pth%
cacls \\%srv%\%ltr%$%pth% /T /C /G %1:C "Domain Admins":F < y.txt > NUL
rmtshare %2 /delete > NUL
rmtshare \\%srv%\%1$=%ltr%:%pth% /Grant %1:Change /Grant "Domain Admins":Full > NUL
netuser -c %1 /D:M: /h:\\%srv%\%1$ > NUL
echo Complete!
goto CLNUP

:ADDCHECK
IF exist \\%2\%1$ echo %1 already has a Homeshare defined on %2
IF exist \\%2\%1$ GOTO CLNUP
set drv=D
IF '%3'=='' goto CREATEPATH
IF not exist %3 echo %3 is not a valid source location
IF not exist %3 GOTO END

:CREATEPATH
IF not exist \\%2\%drv%$\home GOTO OTHERDRV
md \\%2\%drv%$\home\%1
IF ErrorLevel 1 Echo Aborting...
IF ErrorLevel 1 GOTO :CLNUP
cacls \\%2\%drv%$\home\%1 /T /C /G %1:C "Domain Admins":F < y.txt > NUL

:CREATESHR
rmtshare \\%2\%1$=%drv%:\home\%1 /Grant %1:Change /Grant "Domain Admins":Full > NUL
netuser -c %1 /D:M: /h:\\%2\%1$ > NUL
IF not '%3'=='' goto MOVEDATA
echo Complete!
GOTO CLNUP

:OTHERDRV
if '%drv%'=='G' md \\%2\d$\home
if '%drv%'=='G' cacls \\%2\d$\home /C /G "Domain Users":R "Domain Admins":F < y.txt > NUL
if '%drv%'=='G' GOTO SETVAR
if '%drv%'=='F' set drv=G
if '%drv%'=='E' set drv=F
if '%drv%'=='D' set drv=E
GOTO CREATEPATH

:MOVEDATA
echo Moving files and folders to \\%2\%1$
scopy %3 \\%2\%1$ /o/s > NUL
echo Please remove %3 after verifing file integrity
goto CLNUP

:CLNUP
set srv=
set shr=
set ltr=
set pth=
set drv=
IF EXIST tmp.txt erase tmp.txt
IF EXIST tmp2.txt erase tmp2.txt

:END
