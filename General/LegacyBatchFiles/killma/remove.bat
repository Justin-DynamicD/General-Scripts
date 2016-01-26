@echo off
set dom=
set deloldms=
if exist total.txt del total.txt
if exist working.txt del working.txt
FOR /F "SKIP=2 tokens=1,2,3" %%A IN (OUTPUT.TXT) DO echo %%A %%B %%C>>working.txt
type working.txt|find " " /c>total.txt
for /f "tokens=1" %%A in (total.txt) do set deloldms=%%A
cls
echo.
Echo NOTICE: %deloldms% machine accounts found in OUTPUT.TXT, ready for
deletion
Echo Press Ctrl + C to abort or..
echo.
pause
FOR /f "tokens=6" %%a in (output.txt) do set dom=%%a
if "%dom%"=="" goto nodomain
FOR /F "SKIP=2 TOKENS=3" %%A IN (OUTPUT.TXT) DO CALL BAT2 %%A
if exist total.txt del total.txt
if exist working.txt del working.txt
goto end
:nodomain
Echo Domain Name Missing from OUTPUT.TXT
Echo Re-run Gather.BAT
:end 
