@echo off
set SOURCE=C:\InSync\SmartBid\files\data
set DEST=C:\InSync\SmartBid\files\calls

xcopy "%SOURCE%" "%DEST%" /E /C /Y /H /R
echo Archivos y carpetas copiados correctamente.