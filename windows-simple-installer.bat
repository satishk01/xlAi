@echo off
REM Excel-Ollama AI Plugin - Simple Windows Installer
REM This installer only requires Python runtime (no development tools)

echo Excel-Ollama AI Plugin - Simple Installation
echo =============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python 3.8+ from: https://python.org
    echo IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    echo After installing Python, run this installer again.
    echo.
    pause
    exit /b 1
)

echo Python found. Checking version...
python -c "import sys; print(f'Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"

REM Check Python version
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python 3.8 or later is required
    echo Please update your Python installation
    pause
    exit /b 1
)

echo.
echo Installing required Python packages...
echo This may take a few minutes...
echo.

REM Install required packages one by one for better error handling
echo Installing xlwings...
python -m pip install xlwings
if errorlevel 1 (
    echo ERROR: Failed to install xlwings
    pause
    exit /b 1
)

echo Installing pandas...
python -m pip install pandas
if errorlevel 1 (
    echo ERROR: Failed to install pandas
    pause
    exit /b 1
)

echo Installing numpy...
python -m pip install numpy
if errorlevel 1 (
    echo ERROR: Failed to install numpy
    pause
    exit /b 1
)

echo Installing requests...
python -m pip install requests
if errorlevel 1 (
    echo ERROR: Failed to install requests
    pause
    exit /b 1
)

echo Installing aiohttp...
python -m pip install aiohttp
if errorlevel 1 (
    echo ERROR: Failed to install aiohttp
    pause
    exit /b 1
)

echo Installing scikit-learn...
python -m pip install scikit-learn
if errorlevel 1 (
    echo ERROR: Failed to install scikit-learn
    pause
    exit /b 1
)

echo Installing scipy...
python -m pip install scipy
if errorlevel 1 (
    echo ERROR: Failed to install scipy
    pause
    exit /b 1
)

echo.
echo Setting up plugin directories...

REM Create plugin directory
set PLUGIN_DIR=%APPDATA%\ExcelOllamaPlugin
mkdir "%PLUGIN_DIR%" 2>nul
mkdir "%PLUGIN_DIR%\plugin" 2>nul
mkdir "%PLUGIN_DIR%\logs" 2>nul
mkdir "%PLUGIN_DIR%\cache" 2>nul

REM Copy plugin files
echo Copying plugin files...
xcopy /E /Y plugin "%PLUGIN_DIR%\plugin\"
if exist config.json copy config.json "%PLUGIN_DIR%\"

REM Install xlwings Excel add-in
echo Installing Excel add-in...
python -c "import xlwings as xw; xw.addin install"
if errorlevel 1 (
    echo WARNING: xlwings add-in installation failed
    echo You may need to install it manually
)

REM Create Excel startup file
echo Setting up Excel integration...
set EXCEL_STARTUP=%APPDATA%\Microsoft\Excel\XLSTART
mkdir "%EXCEL_STARTUP%" 2>nul

REM Create the Excel startup Python file
(
echo # Excel-Ollama AI Plugin Startup
echo import sys
echo import os
echo.
echo # Add plugin directory to Python path
echo plugin_dir = os.path.join(os.getenv('APPDATA'^), 'ExcelOllamaPlugin'^)
echo sys.path.insert(0, plugin_dir^)
echo.
echo try:
echo     from plugin.main import initialize_plugin
echo     plugin_instance = initialize_plugin(^)
echo     print("Excel-Ollama AI Plugin loaded successfully!"^)
echo except ImportError as e:
echo     print(f"Plugin import error: {e}"^)
echo except Exception as e:
echo     print(f"Plugin initialization error: {e}"^)
echo     import traceback
echo     traceback.print_exc(^)
) > "%EXCEL_STARTUP%\ExcelOllamaPlugin.py"

REM Create desktop shortcut for configuration
echo Creating desktop shortcut...
set DESKTOP=%USERPROFILE%\Desktop
(
echo @echo off
echo echo Opening Excel-Ollama AI Plugin Configuration...
echo python -c "from plugin.ui.dialog_forms import ConfigurationDialog; from plugin.utils.config import PluginConfig; dialog = ConfigurationDialog(PluginConfig(^)^); dialog.show(^)"
echo pause
) > "%DESKTOP%\Configure Ollama Plugin.bat"

echo.
echo =============================================
echo Installation completed successfully!
echo =============================================
echo.
echo Plugin installed to: %PLUGIN_DIR%
echo Excel startup file: %EXCEL_STARTUP%\ExcelOllamaPlugin.py
echo.
echo NEXT STEPS:
echo 1. Restart Excel completely
echo 2. Look for "Ollama AI Analysis" tab in the ribbon
echo 3. If tab doesn't appear, check Excel Add-ins settings
echo 4. Click "Configure" to set your Ollama server URL
echo 5. Test the connection
echo.
echo CONFIGURATION:
if exist config.json (
    echo Your Ollama server URL has been pre-configured.
    echo Check the "Configure Ollama Plugin" shortcut on your desktop.
) else (
    echo You'll need to configure your Ollama server URL in Excel.
    echo Default: http://localhost:11434
    echo For EC2: http://YOUR_EC2_IP:11434
)
echo.
echo TROUBLESHOOTING:
echo - If ribbon tab doesn't appear, restart Excel
echo - Check Excel Add-ins: File ^> Options ^> Add-ins
echo - Enable xlwings add-in if disabled
echo - Check logs in: %PLUGIN_DIR%\logs\
echo.
pause