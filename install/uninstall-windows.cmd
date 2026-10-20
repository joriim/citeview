@echo off
rem Removes the CiteView registration from Word for the current user.
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v CiteView /f >nul 2>&1
rmdir /s /q "%LOCALAPPDATA%\CiteView" 2>nul
echo CiteView removed. Restart Word.
pause
