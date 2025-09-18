@echo off
echo Excel-Ollama AI Plugin - Quick Fix
echo ==================================
echo.

echo This will help fix the "no functions showing" issue.
echo.

REM Get EC2 IP
set /p EC2_IP="Enter your EC2 Public IP address: "

if "%EC2_IP%"=="" (
    echo ERROR: No IP address provided
    pause
    exit /b 1
)

echo.
echo Creating corrected VBA code with your EC2 IP...

REM Create corrected VBA file
powershell -Command "(Get-Content 'corrected-excel-addin.vba') -replace 'YOUR_EC2_IP', '%EC2_IP%' | Set-Content 'FixedExcelOllamaPlugin.vba'"

echo ✅ Fixed VBA code created: FixedExcelOllamaPlugin.vba
echo.

echo MANUAL STEPS TO FIX:
echo ====================
echo.
echo 1. Open Excel
echo 2. Press Alt+F11 (VBA Editor)
echo 3. Find your existing add-in project
echo 4. Delete the old Module1 (right-click → Remove)
echo 5. Insert → Module (create new module)
echo 6. Copy ALL content from FixedExcelOllamaPlugin.vba
echo 7. Paste into the new module
echo 8. Press Ctrl+S to save
echo 9. Close VBA Editor (Alt+Q)
echo 10. Press Alt+F8 - you should now see the functions!
echo.

echo ALTERNATIVE - RECREATE ADD-IN:
echo =============================
echo.
echo 1. File → Options → Add-ins → Uncheck current add-in
echo 2. Create new workbook
echo 3. Press Alt+F11
echo 4. Insert → Module
echo 5. Copy content from FixedExcelOllamaPlugin.vba
echo 6. Save as Excel Add-in (.xlam)
echo 7. Install the new add-in
echo.

echo FUNCTIONS YOU SHOULD SEE:
echo =========================
echo • AnalyzeSelectedData
echo • AskQuestionAboutData  
echo • GenerateComprehensiveReport
echo • AnalyzeTrends
echo • DetectPatterns
echo • ConfigureOllamaServer
echo • TestOllamaConnection
echo • ShowHelp
echo • CreateSampleData
echo.

echo 💡 TIP: Try CreateSampleData first to test with sample data!
echo.

pause