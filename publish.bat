@echo on
dotnet publish SmartBid.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o C:\InSync\SmartBid\publish

set SOURCE=C:\InSync\SmartBid\publish
set DEST=C:\InSync\test\SmartBid

xcopy "C:\InSync\SmartBid\properties.xml" "C:\InSync\SmartBid\publish\properties.xml" /C /Y /H /R

xcopy "%SOURCE%\*" "%DEST%" /C /Y /H /R
echo Archivos copiados correctamente.

timeout /t 3 >nul
