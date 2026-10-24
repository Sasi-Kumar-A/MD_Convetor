@echo off
setlocal
cd /d "%~dp0"

if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate.bat

echo Checking and installing dependencies...
pip install -r requirements.txt

echo Starting MD Converter Studio...
python app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with an error.
    pause
)
