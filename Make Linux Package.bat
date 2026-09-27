@echo off
title Make Linux Package
cd /d "%~dp0"
echo Copying the latest files into "LINUX VERSION" and making the zip...
echo.
python "%~dp0make_linux_package.py"
if errorlevel 1 (
    echo.
    echo *** FAILED - see the messages above. ***
) else (
    echo.
    echo Done. Send "JJs Setlist - Linux.zip" to your Linux user,
    echo or publish the "LINUX VERSION" folder as a repository.
)
echo.
pause
