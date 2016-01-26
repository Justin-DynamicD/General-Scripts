@echo off
Echo Save the batch file "AU_Clean_SID.cmd". This batch file will do the following:
Echo 1.    Stop the wuauserv service
Echo 2.    Delete the AccountDomainSid registry key (if it exists)
Echo 3.    Delete the PingID registry key (if it exists)
Echo 4.    Delete the SusClientId registry key (if it exists)
Echo 5.    Restart the wuauserv service
Echo 6.    Resets the Authorization Cookie

pause
@echo on
net stop wuauserv
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v AccountDomainSid /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v PingID /f
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f
net start wuauserv
wuauclt /resetauthorization /detectnow


