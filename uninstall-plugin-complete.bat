@echo off
REM ============================================================================
REM Complete Uninstall Script for Excel-Ollama AI Plugin
REM Removes all traces of the plugin from your system
REM ============================================================================

setlocal enabledelayedexpansion
color 0C

echo.
echo ========================================================================
echo                    COMPLETE PLUGIN UNINSTALL
echo ========================================================================
echo.
echo This will completely remove the Excel-Ollama AI Plugin from your system.
echo.
echo What will be removed:
echo - Excel Add-in files (.xlam)
echo - Registry entries (if any)
echo - Temporary files
echo - Configuration files
echo.

choice /C YN /M "Are you sure you want to completely uninstall the plugin? (Y/N)"
if errorlevel 2 goto :EOF

echo.
echo ========================================================================
echo                           UNINSTALL STEPS
echo ========================================================================
echo.

REM Step 1: Close Excel if running
echo Step 1: Checking for running Excel processes...
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo WARNING: Excel is currently running
    echo Please close Excel before continuing
    echo.
    choice /C YN /M "Force close Excel processes? (Y/N)"
    if errorlevel 1 (
        echo Closing Excel processes...
        taskkill /F /IM EXCEL.EXE >nul 2>&1
        timeout /t 3 >nul
        echo Excel processes closed
    )
)
echo.

REM Step 2: Remove add-in files
echo Step 2: Removing add-in files...

REM Common add-in locations
set "ADDINS_DIR=%APPDATA%\Microsoft\AddIns"
set "CURRENT_DIR=%~dp0"

REM Remove from Excel Add-ins directory
if exist "%ADDINS_DIR%\OllamaAI_Enterprise.xlam" (
    del "%ADDINS_DIR%\OllamaAI_Enterprise.xlam" >nul 2>&1
    if !errorlevel! equ 0 (
        echo   - Removed: %ADDINS_DIR%\OllamaAI_Enterprise.xlam
    ) else (
        echo   - Could not remove: %ADDINS_DIR%\OllamaAI_Enterprise.xlam
    )
)

REM Remove from current directory
if exist "%CURRENT_DIR%OllamaAI_Enterprise.xlam" (
    del "%CURRENT_DIR%OllamaAI_Enterprise.xlam" >nul 2>&1
    if !errorlevel! equ 0 (
        echo   - Removed: %CURRENT_DIR%OllamaAI_Enterprise.xlam
    )
)

REM Remove other variations
for %%f in (
    "enterprise-excel-addin.xlam"
    "bulletproof-excel-addin.xlam"
    "fixed-excel-addin.xlam"
    "excel-ollama-plugin.xlam"
) do (
    if exist "%ADDINS_DIR%\%%f" (
        del "%ADDINS_DIR%\%%f" >nul 2>&1
        echo   - Removed: %%f
    )
    if exist "%CURRENT_DIR%%%f" (
        del "%CURRENT_DIR%%%f" >nul 2>&1
        echo   - Removed: %%f from current directory
    )
)

echo.

REM Step 3: Remove temporary files
echo Step 3: Removing temporary files...

for %%f in (
    "create_enterprise_addin.ps1"
    "create_enterprise_addin_fixed.ps1"
    "temp_addin.xlsm"
    "ENTERPRISE_SETUP_INSTRUCTIONS.txt"
) do (
    if exist "%CURRENT_DIR%%%f" (
        del "%CURRENT_DIR%%%f" >nul 2>&1
        echo   - Removed: %%f
    )
)

echo.

REM Step 4: Registry cleanup (optional)
echo Step 4: Registry cleanup...
echo.
echo The plugin may have created registry entries for Excel add-ins.
echo These are usually automatically cleaned up by Excel.
echo.
choice /C YN /M "Do you want to check for registry entries? (Y/N)"
if errorlevel 1 (
    echo Checking registry...
    
    REM Check for add-in registry entries
    reg query "HKCU\Software\Microsoft\Office\Excel\Addins" /s 2>nul | find /i "ollama" >nul
    if !errorlevel! equ 0 (
        echo   - Found Ollama-related registry entries
        echo   - These will be cleaned up automatically by Excel
    ) else (
        echo   - No Ollama-related registry entries found
    )
)

echo.

REM Step 5: Excel add-in list cleanup instructions
echo Step 5: Excel Add-in List Cleanup Instructions...
echo.
echo IMPORTANT: Manual steps required in Excel:
echo.
echo 1. Open Microsoft Excel
echo 2. Go to File ^> Options ^> Add-ins
echo 3. At the bottom, select "Excel Add-ins" and click "Go..."
echo 4. Look for any Ollama-related add-ins in the list
echo 5. Uncheck any Ollama add-ins
echo 6. If you see "OllamaAI_Enterprise" with a missing file error:
echo    - Select it and click "Remove" or "Delete"
echo 7. Click OK
echo.

REM Step 6: VBA Project cleanup
echo Step 6: VBA Project Cleanup Instructions...
echo.
echo If you manually added VBA code to a workbook:
echo.
echo 1. Open the workbook with the VBA code
echo 2. Press Alt+F11 to open VBA Editor
echo 3. In Project Explorer, find modules with Ollama code
echo 4. Right-click the module and select "Remove [ModuleName]"
echo 5. Choose "No" when asked to export (unless you want to keep a backup)
echo 6. Save the workbook
echo.

REM Step 7: Verification
echo Step 7: Verification...
echo.
echo To verify complete removal:
echo.
echo 1. Restart Excel
echo 2. You should NOT see the "Enterprise Excel-Ollama AI Plugin loaded!" message
echo 3. Press Alt+F8 - you should NOT see any Ollama functions
echo 4. Go to File ^> Options ^> Add-ins - no Ollama add-ins should be listed
echo.

echo ========================================================================
echo                           UNINSTALL COMPLETED
echo ========================================================================
echo.
echo FILES REMOVED:
echo - Excel Add-in files (.xlam)
echo - Temporary PowerShell scripts
echo - Setup instruction files
echo.
echo MANUAL STEPS REQUIRED:
echo 1. Remove add-in from Excel's add-in list (see instructions above)
echo 2. Remove VBA code from any workbooks (if manually added)
echo 3. Restart Excel to verify complete removal
echo.
echo VERIFICATION:
echo - No "plugin loaded" message on Excel startup
echo - No Ollama functions in Alt+F8 macro list
echo - No Ollama add-ins in Excel Options ^> Add-ins
echo.
echo If you see any remaining traces:
echo 1. Check Excel Options ^> Add-ins and remove manually
echo 2. Check for .xlam files in: %ADDINS_DIR%
echo 3. Restart Excel after manual cleanup
echo.
echo ========================================================================

echo.
echo Press any key to exit...
pause >nul