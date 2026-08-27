@echo off
REM ---------------------------------------------------------------------------
REM DOMIN8 Omnichannel Reports - web app
REM Double-click this. It installs what is missing, starts the app, and opens
REM your browser. Close this window to stop the app.
REM ---------------------------------------------------------------------------
title DOMIN8 Reports
cd /d "%~dp0"

python -c "import streamlit" 2>nul
if errorlevel 1 (
    echo First run - installing dependencies, this takes a minute...
    python -m pip install -q -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Install failed. Check that Python is on your PATH.
        pause
        exit /b 1
    )
)

echo Starting DOMIN8 Reports...
echo.
echo   The app will open at http://localhost:8501
echo   Leave this window open. Close it to stop the app.
echo.
start "" http://localhost:8501
python -m streamlit run app.py --server.port 8501 --server.headless true
pause
