@echo off
Title THREDrDB Uninstall

cls
echo Uninstalling THREDrDB...
echo.
set /p answer=Do you want to continue (Y/N)? 
if /i "%answer:~,1%" EQU "Y" goto UninstallIt
if /i "%answer:~,1%" EQU "N" exit /b
echo Please type Y for Yes or N for No
goto :eof
:UninstallIt
echo Unregistering THREDrDB add-in from Excel...
echo Set objExcel = CreateObject("Excel.Application") > "%TEMP%\UnregisterXLL.vbs"
echo objExcel.Visible = False >> "%TEMP%\UnregisterXLL.vbs"
echo objExcel.Workbooks.Add >> "%TEMP%\UnregisterXLL.vbs"
echo For Each addIn In objExcel.AddIns >> "%TEMP%\UnregisterXLL.vbs"
echo     If InStr(LCase(addIn.FullName), LCase("ThredrDB_add-in-AddIn64-packed.xll")) > 0 Then addIn.Installed = False >> "%TEMP%\UnregisterXLL.vbs"
echo Next >> "%TEMP%\UnregisterXLL.vbs"
echo objExcel.Quit >> "%TEMP%\UnregisterXLL.vbs"
cscript "%TEMP%\UnregisterXLL.vbs" > nul 2>&1
del "%TEMP%\UnregisterXLL.vbs" > nul 2>&1
echo Removing THREDrDB files...
if exist "%APPDATA%\ThredrDB" rmdir /s /q "%APPDATA%\ThredrDB"
echo Removing Add/Remove Programs entry...
reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /f > nul 2>&1
echo Uninstallation complete! Please restart Excel to ensure THREDrDB is no longer loaded.
echo This window will close in 5 seconds...
timeout 5 > nul
exit /b
:Error
echo An error occurred during uninstallation. Please try again or contact support@thredrdb.com.
echo This window will close in 10 seconds...
timeout 10 > nul
exit /b 1