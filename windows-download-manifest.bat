@echo off
REM Excel-Ollama AI Plugin - Windows Manifest Downloader
REM Downloads the Excel add-in manifest from EC2 server

echo Excel-Ollama AI Plugin - Manifest Downloader
echo ===============================================
echo.

REM Get EC2 server URL from user
set /p EC2_IP="Enter your EC2 Public IP address: "

if "%EC2_IP%"=="" (
    echo ERROR: No IP address provided
    pause
    exit /b 1
)

set SERVER_URL=http://%EC2_IP%:3000
set MANIFEST_URL=%SERVER_URL%/manifest.xml

echo.
echo Testing connection to server...
echo Server URL: %SERVER_URL%

REM Test if server is reachable
curl -s --connect-timeout 10 %SERVER_URL% >nul 2>&1
if errorlevel 1 (
    echo ERROR: Cannot connect to server at %SERVER_URL%
    echo.
    echo Please check:
    echo 1. EC2 instance is running
    echo 2. Security Group allows port 3000 from your IP
    echo 3. Web server is running on EC2
    echo.
    pause
    exit /b 1
)

echo SUCCESS: Server is reachable!
echo.

REM Create downloads directory
if not exist "%USERPROFILE%\Downloads\ExcelOllamaPlugin" (
    mkdir "%USERPROFILE%\Downloads\ExcelOllamaPlugin"
)

set DOWNLOAD_DIR=%USERPROFILE%\Downloads\ExcelOllamaPlugin

echo Downloading manifest file...
curl -o "%DOWNLOAD_DIR%\manifest.xml" %MANIFEST_URL%

if errorlevel 1 (
    echo ERROR: Failed to download manifest file
    pause
    exit /b 1
)

echo.
echo ===============================================
echo Download completed successfully!
echo ===============================================
echo.
echo Manifest file saved to:
echo %DOWNLOAD_DIR%\manifest.xml
echo.
echo NEXT STEPS:
echo 1. Open Microsoft Excel
echo 2. Go to Insert ^> Get Add-ins
echo 3. Click "Upload My Add-in"
echo 4. Select the downloaded manifest.xml file
echo 5. Click "Upload"
echo.
echo The add-in will appear in Excel's task pane.
echo Configure it with server URL: %SERVER_URL%
echo.
echo ALTERNATIVE METHOD:
echo You can also use the direct URL in Excel:
echo %MANIFEST_URL%
echo.

REM Open the download folder
echo Opening download folder...
explorer "%DOWNLOAD_DIR%"

echo.
pause