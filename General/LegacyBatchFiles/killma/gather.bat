@echo off
if "%1"=="" goto nodomain
set dom=%1
set deloldms=
netdom /d:%1 bdc | find "Found PDC" > temp4.txt
for /f "tokens=3" %%a in (temp4.txt) do set pdc=%%a
if exist ms2.txt del ms2.txt
if exist output.txt del output.txt
if exist out2.txt del out2.txt
if exist temp4.txt del temp4.txt
echo.
echo Generating Server List of Member Servers and Workstations
echo.
echo Please Wait...
netdom /d:%1 /noverbose member > MS.TXT
for /F "delims=\\ tokens=1" %%a in (ms.txt) do echo %%a >> MS2.TXT
cls
echo.
echo Generating List of Member Servers and Workstations - Done
echo.
echo List Generated.  Checking Password Ages.
echo.
echo Please Wait...
for /F "tokens=1" %%a in (ms2.txt) do call oldms2.bat %%a %dom% %pdc%
sort < output.txt > out2.txt
del output.txt
echo Machine account ages for domain: %dom% >> output.txt
echo ------------------------------------------------ >> output.txt
type out2.txt >> output.txt
if exist ms.txt del ms.txt
if exist out2.txt del out2.txt
if exist temp3.txt del temp3.txt
if exist ms2.txt del ms2.txt
if exist temp.txt del temp.txt
if exist temp4.txt del temp4.txt
if exist total.txt del total.txt
if exist working.txt del working.txt
FOR /F "SKIP=2 tokens=1,2,3" %%A IN (OUTPUT.TXT) DO echo %%A %%B %%C>>working.txt
type working.txt|find " " /c>total.txt
for /f "tokens=1" %%A in (total.txt) do set deloldms=%%A
echo.
Echo List Complete
echo.
Echo %deloldms% machine accounts found.
echo.
echo Now edit OUTPUT.TXT and remove all valid machine accounts.
echo Machine accounts remaining in OUTPUT.TXT will be deleted.
echo After OUTPUT.TXT has been modified, run DELETE.BAT to
echo delete machine accounts.
echo.
if exist total.txt del total.txt
if exist working.txt del working.txt
goto end
:nodomain
echo Specify the target domain on the command line
echo EXAMPLE: gather MyDomainName
:end 
