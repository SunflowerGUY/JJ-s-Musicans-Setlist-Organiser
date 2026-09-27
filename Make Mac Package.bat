@echo off
title Make Mac Package
cd /d "%~dp0"
echo Copying the latest files into "APPLE MAC VERSION" and making the zip...
echo.
python "%~dp0make_mac_package.py"
if errorlevel 1 (
    echo.
    echo *** FAILED - see the messages above. ***
) else (
    echo.
    echo Done. Send "JJs Setlist - Mac.zip" to your Mac user.
)
echo.
pause
