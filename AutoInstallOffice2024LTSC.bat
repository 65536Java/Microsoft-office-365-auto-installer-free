@ECHO OFF
powershell -Command "Invoke-WebRequest -Uri 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18730-20142.exe' -OutFile '%CD%\od.exe'"
powershell -Command "Invoke-WebRequest -Uri 'https://drive.usercontent.google.com/download?id=1AEy4FwHK_818ohzEMLBKrrf3jeDW5You&export=download&authuser=0&confirm=t&uuid=064dda4e-ae46-4dd7-86de-bf89b8f2f26e&at=ALoNOgm27Snk0rXjHj7BjtNIUreJ:1749555888635' -OutFile '%CD%/cfg.xml'"
od
del configuration-Office365-x64.xml
echo Downloading office 2024...
setup /download cfg.xml
setup /configure cfg.xml
echo Activating office 2024...
cd C:\Program Files\Microsoft Office\Office16
cscript ospp.vbs /sethst:kms.03k.org 
cscript ospp.vbs /act
pause