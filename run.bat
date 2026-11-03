@echo off

rem Check if python has all required libraries already installed
python -c "import customtkinter, pandas, openpyxl, PIL" >nul 2>&1
if %errorlevel% equ 0 (
    start "" pythonw "%~dp0EOT_GUI.py"
    exit /b
)

rem Check if py launcher has all required libraries already installed
py -c "import customtkinter, pandas, openpyxl, PIL" >nul 2>&1
if %errorlevel% equ 0 (
    start "" pyw "%~dp0EOT_GUI.py"
    exit /b
)

rem If packages are missing, show console and install them
echo ==========================================
echo Setting up P6 EOT Automation Toolkit...
echo ==========================================

python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Installing missing packages...
    pip install -r "%~dp0requirements.txt"
    start "" pythonw "%~dp0EOT_GUI.py"
    exit /b
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Installing missing packages...
    py -m pip install -r "%~dp0requirements.txt"
    start "" pyw "%~dp0EOT_GUI.py"
    exit /b
)

rem Final fallback to Windows file association
start "" "%~dp0EOT_GUI.py"
