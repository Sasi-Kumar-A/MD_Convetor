@echo off
title MD Converter Studio
echo Starting MD Converter Studio...
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with an error.
    pause
)
