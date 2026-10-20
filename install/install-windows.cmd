@echo off
rem CiteView installer for Word on Windows (Microsoft 365 / Word 2019+).
rem Downloads the add-in manifest and registers it for the current user. No admin rights needed.
setlocal
set "DIR=%LOCALAPPDATA%\CiteView"
set "URL=https://joriim.github.io/citeview/manifest-github.xml"
if not exist "%DIR%" mkdir "%DIR%"
echo Downloading manifest...
powershell -NoProfile -Command "Invoke-WebRequest -UseBasicParsing '%URL%' -OutFile '%DIR%\manifest.xml'" || goto :fail
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v CiteView /t REG_SZ /d "%DIR%\manifest.xml" /f >nul || goto :fail
echo.
echo CiteView installed. Close Word completely, open it again, then use References ^> Show sources.
pause
exit /b 0
:fail
echo Installation failed. See https://joriim.github.io/citeview for the manual method.
pause
exit /b 1
