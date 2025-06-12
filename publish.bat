@echo on
dotnet publish SmartBid.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true -o C:\InSync\SmartBid\publish


set SOURCE=C:\InSync\SmartBid\publish
set DEST=C:\InSync\test\SmartBid


xcopy "%SOURCE%\*" "%DEST%" /C /Y /H /R
echo Archivos copiados correctamente.
