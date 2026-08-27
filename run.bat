@echo off
REM ---------------------------------------------------------------------------
REM DOMIN8 omnichannel report.
REM
REM   run.bat              build from the files already in reports\input\uniware
REM   run.bat --fetch      pull fresh Uniware data first (asks for the password)
REM   run.bat --status     show what is on disk
REM
REM Any run_pipeline.py flag can be passed straight through.
REM ---------------------------------------------------------------------------
cd /d "%~dp0"

echo %* | findstr /C:"--fetch" >nul
if %errorlevel%==0 (
    if "%UNIWARE_USER%"=="" set UNIWARE_USER=domin8@abstrabit.com
    if "%UNIWARE_PASS%"=="" set /p UNIWARE_PASS=Uniware password:
)

python run_pipeline.py %*
echo.
pause
