@echo off
chcp 65001

copy "CPI Sheets Updater.dll"  "%APPDATA%\Autodesk\Revit\Addins\2019"
copy "CPI Sheets Updater.addin"  "%APPDATA%\Autodesk\Revit\Addins\2019"

copy "CPI Sheets Updater.dll"  "%APPDATA%\Autodesk\Revit\Addins\2022"
copy "CPI Sheets Updater.addin"  "%APPDATA%\Autodesk\Revit\Addins\2022"

cls

echo Revit 2019 OK
echo Revit 2022 OK

@echo off
echo Готово!
echo Для выхода нажмите любую клавишу . . .
pause>nul
