@echo off
TITLE A&H Brass Shopify Feed Generator

echo A&H Brass Shopify Feed Generator
echo --------------------------------

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: Python is required but not installed.
    pause
    exit /b 1
)

REM Check if requirements.txt exists
if not exist requirements.txt (
    echo Error: requirements.txt not found.
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

REM Check if app.py exists
if not exist app.py (
    echo Error: app.py not found.
    pause
    exit /b 1
)

REM Run the app
echo Starting Streamlit app...
streamlit run app.py

echo App closed.
pause 