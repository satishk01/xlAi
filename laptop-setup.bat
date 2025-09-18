@echo off
echo Excel-Ollama AI Plugin - Laptop Setup
echo =====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python found. Checking version...
python -c "import sys; print(f'Python {sys.version}')"
echo.

REM Check if Excel is available
echo Checking for Microsoft Excel...
if exist "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE" (
    echo Excel found: Office 2016/2019/2021
) else if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE" (
    echo Excel found: Office 2016/2019/2021 (32-bit)
) else if exist "%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE" (
    echo Excel found: Office 2016
) else (
    echo WARNING: Excel not found in standard locations
    echo Please ensure Microsoft Excel 2016 or later is installed
)
echo.

REM Install Python dependencies
echo Installing Python dependencies...
echo This may take a few minutes...
python -m pip install --upgrade pip
python -m pip install xlwings pandas numpy requests aiohttp scikit-learn scipy tkinter-tooltip
if errorlevel 1 (
    echo ERROR: Failed to install some dependencies
    echo Trying alternative installation...
    python -m pip install xlwings pandas numpy requests aiohttp scikit-learn scipy
)
echo.

REM Install the plugin
echo Installing Excel-Ollama AI Plugin...
python install.py --install
if errorlevel 1 (
    echo ERROR: Plugin installation failed
    echo Please check the error messages above
    pause
    exit /b 1
)

echo.
echo =====================================
echo Installation completed successfully!
echo =====================================
echo.
echo Next steps:
echo 1. Make sure your EC2 Ollama server is running
echo 2. Open Excel
echo 3. Look for "Ollama AI Analysis" tab in the ribbon
echo 4. Click "Configure" to set your EC2 server URL
echo 5. Test the connection
echo.
echo Your EC2 server URL should be:
echo http://YOUR_EC2_PUBLIC_IP:11434
echo.
echo Need help? Check DEPLOYMENT_GUIDE.md
echo.
pause