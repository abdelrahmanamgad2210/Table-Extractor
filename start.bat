@echo off
cd /d "%~dp0"

:: Check for .env and prompt for key if missing
if not exist ".env" (
    echo.
    echo  ============================================================
    echo   First-time setup: Gemini API Key required
    echo   Get a free key at: https://aistudio.google.com/apikey
    echo  ============================================================
    echo.
    set /p APIKEY=" Paste your API key here: "
    echo GEMINI_API_KEY=%APIKEY%> .env
    echo.
    echo  Key saved to .env — you won't be asked again.
    echo.
)

python server.py
pause
