@echo off
setlocal

echo ==========================================
echo  P6 EOT Automation Toolkit - Web App
echo ==========================================

set "PORT=8000"
set "HOST=127.0.0.1"

rem --- Check for web dependencies ---
python -c "import fastapi, uvicorn, jinja2" >nul 2>&1
if %errorlevel% equ 0 goto :launch_web

echo Installing web dependencies...
pip install fastapi "uvicorn[standard]" jinja2 python-multipart
if %errorlevel% neq 0 (
    echo [ERROR] pip install failed. Please install dependencies manually.
    pause
    exit /b 1
)

:launch_web
echo.
echo Starting web app at http://%HOST%:%PORT%
echo Open your browser to: http://localhost:%PORT%
echo Press Ctrl+C to stop the server.
echo.

python -m uvicorn web_app:app --host %HOST% --port %PORT% --reload
pause

