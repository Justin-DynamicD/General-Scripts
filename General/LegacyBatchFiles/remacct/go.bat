@Echo Off
rem %1 is the backup directory

:GETUSRS
IF NOT EXIST %1 echo %1 is an invalid backup directory
IF NOT EXIST %1 goto END
IF NOT EXIST list.txt echo Please supply accounts in list.txt file
IF NOT EXIST list.txt goto END

:MVEDATA
echo --- Moving User Home Shares ---
FOR /F "TOKENS=1,2* delims=," %%A IN (list.txt) DO CALL sub1.bat %%A %%B %1

:DELUSERS
echo --- Accounts have been backed up, ready to remove ---
pause
echo --- Removing User Accounts ---
FOR /F "TOKENS=1 delims=," %%A IN (list.txt) DO CALL sub2.bat %%A

:END
