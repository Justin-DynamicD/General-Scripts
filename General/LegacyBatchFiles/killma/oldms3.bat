rem %1 = date
rem %2 = time
rem %3 = member server
echo %1 > temp3.txt
for /F "delims=/ tokens=1,2,3" %%a in (temp3.txt) do oldms4.bat %%a %%b %%c %2 %3
