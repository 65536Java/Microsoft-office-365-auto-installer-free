@ECHO OFF
powershell -Command "Invoke-WebRequest -Uri 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18730-20142.exe' -OutFile '%CD%\od.exe'"
powershell -Command "Invoke-WebRequest -Uri 'https://download1507.mediafire.com/ih4pdg9cruzg3V8Lp3ymS9_aRi7s08KWgZTWiFY9KyqZS2VqO9ikS-fLVwYLCwXS41zHs-4t8iS3mltUTcsyHNBjvGYkZoeQiwvShmkcw22GShQzQzTqlKJhTgDIBY88datsui0-92Z_OjARVDztJ8BQzZli5aG_CaCkGgiKo1x_TQ/2t03qf9nv8agxh7/config.xml' -OutFile '%CD%/config.xml'"
od
del configuration-Office365-x64.xml
echo Downloading office 2024...
setup /download config.xml
setup /configure config.xml
echo Activating office 2024...
cd C:\Program Files\Microsoft Office\Office16
cscript ospp.vbs /sethst:kms.03k.org 
cscript ospp.vbs /act
pause
