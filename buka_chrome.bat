@echo off
REM Buka Chrome khusus bot remark NTB (port remote debugging + profil terpisah). Login SSO manual.
REM   buka_chrome.bat          -> port 9340
REM   buka_chrome.bat 9341     -> port lain (untuk jalan paralel, login lagi di jendela itu)
set PORT=9340
if not "%1"=="" set PORT=%1
set PROFIL=%~dp0.profil_chrome_%PORT%

set CHROME=
for %%P in ("%ProgramFiles%\Google\Chrome\Application\chrome.exe" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" "%LocalAppData%\Google\Chrome\Application\chrome.exe") do if not defined CHROME if exist %%P set CHROME=%%~P
if not defined CHROME goto tanpa_chrome

start "" "%CHROME%" --remote-debugging-port=%PORT% --user-data-dir="%PROFIL%" --no-first-run --no-default-browser-check "https://fasih-sm.bps.go.id/app/"
echo Chrome dibuka di port %PORT%. Login SSO FASIH di jendela itu, lalu jalankan bot (lihat README.md).
exit /b 0

:tanpa_chrome
echo !! chrome.exe tidak ditemukan. Pasang Google Chrome dulu.
exit /b 1
