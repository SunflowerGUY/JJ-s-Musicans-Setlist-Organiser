@echo off
setlocal
title Build JJ's Setlist
cd /d "%~dp0"

set "APPNAME=JJs Setlist"
set "WORK=%TEMP%\setlist_build"

echo ==================================================
echo   JJ's Setlist  -  build "%APPNAME%.exe"
echo ==================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install Python from python.org and tick
    echo "Add Python to PATH", then run this again.
    goto :fail
)

rem --- Close the app if it's running (it can't be replaced while open) ---
call :isrunning
if errorlevel 1 goto :notrunning
echo "%APPNAME%.exe" is running - closing it...
echo (If it asks to save your setlist, answer within 20 seconds.)
taskkill /im "%APPNAME%.exe" >nul 2>nul
set /a WAITED=0
:waitclose
rem ping is used as a 1-second pause (timeout fails without a keyboard)
ping -n 2 127.0.0.1 >nul
call :isrunning
if errorlevel 1 goto :closed
set /a WAITED+=1
if %WAITED% lss 20 goto :waitclose
echo Still running - forcing it to close...
taskkill /f /t /im "%APPNAME%.exe" >nul 2>nul
ping -n 3 127.0.0.1 >nul
call :isrunning
if not errorlevel 1 (
    echo Could not close "%APPNAME%.exe". Please close it and try again.
    goto :fail
)
:closed
echo       Closed.
echo.
:notrunning

echo [1/3] Checking build tools...
python -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo       Installing PyInstaller...
    python -m pip install --upgrade pyinstaller
    if errorlevel 1 goto :fail
)
python -c "import PIL" >nul 2>nul
if errorlevel 1 (
    echo       Installing Pillow...
    python -m pip install pillow
    if errorlevel 1 goto :fail
)
python -c "import certifi" >nul 2>nul
if errorlevel 1 (
    echo       Installing certifi...
    python -m pip install certifi
    if errorlevel 1 goto :fail
)
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo       Installing openpyxl...
    python -m pip install openpyxl
    if errorlevel 1 goto :fail
)

echo [2/3] Compiling (this takes a minute or two)...
if exist "%~dp0ARTWORK\JJ SETLIST Icon Source.png" (
    python "%~dp0make_icons.py"
    if errorlevel 1 goto :fail
)
set "ICONOPT="
if exist "%~dp0setlist.ico" set ICONOPT=--icon "%~dp0setlist.ico" --add-data "%~dp0setlist.ico;."
if exist "%~dp0setlist.ico" (echo       Using icon: setlist.ico) else (echo       No setlist.ico found - using the default icon.)
python -m PyInstaller --noconfirm --clean --log-level WARN ^
    --onefile --windowed --name "%APPNAME%" ^
    --hidden-import openpyxl --hidden-import certifi %ICONOPT% ^
    --distpath "%~dp0." --workpath "%WORK%" --specpath "%WORK%" ^
    "%~dp0jjs_setlist.py"
if errorlevel 1 goto :fail

echo [3/3] Cleaning up...
rmdir /s /q "%WORK%" 2>nul

echo.
echo ==================================================
echo   Done!  "%APPNAME%.exe" is in:
echo   %~dp0
echo.
echo   Keep it in this folder so it finds your song
echo   database, saved setlists and settings.
echo ==================================================
echo.
pause
exit /b 0

:isrunning
rem errorlevel 0 = running, 1 = not running
tasklist /fi "imagename eq %APPNAME%.exe" | find /i "%APPNAME%.exe" >nul
exit /b %errorlevel%

:fail
echo.
echo *** Build FAILED - see the messages above. ***
echo.
pause
exit /b 1
