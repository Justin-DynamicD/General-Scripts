rem %1 = member server
rem %2 = domain
rem %3 = pdc
nltest /server:%3 /user:%1$ | find "PasswordLastSet" > temp.txt
for /F "delims== tokens=2" %%a in (temp.txt) do oldms3.bat %%a %1 
