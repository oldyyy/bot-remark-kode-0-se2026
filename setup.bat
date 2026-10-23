@echo off
REM Persiapan PC (sekali saja): pasang dependensi bot.
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY goto tanpa_python
%PY% --version
%PY% -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto gagal
if not exist input mkdir input
echo.
echo Selesai. Langkah berikutnya:
echo   1. buka_chrome.bat  lalu login SSO FASIH di jendela Chrome itu
echo   2. %PY% bot.py --dry-run --limit 3
pause
exit /b 0
:tanpa_python
echo !! Python belum terpasang. Pasang Python 3.10+ dari python.org (centang "Add to PATH").
pause
exit /b 1
:gagal
echo !! Pemasangan dependensi gagal (cek koneksi internet).
pause
exit /b 1
