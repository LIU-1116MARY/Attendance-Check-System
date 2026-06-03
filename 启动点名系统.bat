@echo off
title Roll Call System
echo ================================
echo     Classroom Roll Call System
echo ================================
echo.
python "%~dp0roll_call.py"
if errorlevel 1 (
    echo.
    echo Error: Please make sure Python is installed.
    echo Press any key to exit...
    pause >nul
)