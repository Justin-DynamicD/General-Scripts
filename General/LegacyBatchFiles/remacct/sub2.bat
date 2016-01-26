REM %1 = Account

echo Removing %1
net user %1 /DELETE /DOMAIN
