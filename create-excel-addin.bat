@echo off
REM Excel-Ollama AI Plugin - Add-in Creator
REM Creates .xlam add-in file for Excel

echo Excel-Ollama AI Plugin - Add-in Creator
echo ========================================
echo.

REM Get EC2 IP from user
set /p EC2_IP="Enter your EC2 Public IP address: "

if "%EC2_IP%"=="" (
    echo ERROR: No IP address provided
    pause
    exit /b 1
)

echo.
echo Creating Excel Add-in with server: http://%EC2_IP%:11434
echo.

REM Create VBA file with actual IP
powershell -Command "(Get-Content 'create-excel-addin.vba') -replace 'YOUR_EC2_IP', '%EC2_IP%' | Set-Content 'ExcelOllamaAIPlugin.vba'"

if not exist "ExcelOllamaAIPlugin.vba" (
    echo ERROR: Could not create VBA file
    pause
    exit /b 1
)

echo ✅ VBA code created: ExcelOllamaAIPlugin.vba
echo.

REM Create VBScript to build the add-in
echo Creating add-in builder script...

(
echo Dim xl, wb, vbComp, fso, file
echo Dim vbaCode
echo.
echo ' Create Excel application
echo Set xl = CreateObject("Excel.Application"^)
echo xl.Visible = True
echo xl.DisplayAlerts = False
echo.
echo ' Create new workbook
echo Set wb = xl.Workbooks.Add
echo.
echo ' Read VBA code from file
echo Set fso = CreateObject("Scripting.FileSystemObject"^)
echo Set file = fso.OpenTextFile("ExcelOllamaAIPlugin.vba", 1^)
echo vbaCode = file.ReadAll
echo file.Close
echo.
echo ' Add VBA module
echo Set vbComp = wb.VBProject.VBComponents.Add(1^)
echo vbComp.CodeModule.AddFromString vbaCode
echo.
echo ' Save as Excel Add-in
echo wb.SaveAs "ExcelOllamaAIPlugin.xlam", 18
echo.
echo ' Show completion message
echo MsgBox "Excel Add-in created successfully!" ^& vbCrLf ^& vbCrLf ^& _
echo        "File: ExcelOllamaAIPlugin.xlam" ^& vbCrLf ^& _
echo        "Server: http://%EC2_IP%:11434" ^& vbCrLf ^& vbCrLf ^& _
echo        "Install it manually in Excel: File ^> Options ^> Add-ins"
echo.
echo ' Close workbook
echo wb.Close False
echo.
echo ' Cleanup
echo xl.DisplayAlerts = True
echo xl.Quit
echo Set xl = Nothing
) > build-addin.vbs

echo ✅ Add-in builder created: build-addin.vbs
echo.

REM Run the VBScript to create the add-in
echo Creating Excel Add-in (.xlam file)...
echo This will open Excel briefly...
echo.

cscript //nologo build-addin.vbs

if exist "ExcelOllamaAIPlugin.xlam" (
    echo.
    echo ========================================
    echo ✅ SUCCESS! Excel Add-in Created
    echo ========================================
    echo.
    echo File created: ExcelOllamaAIPlugin.xlam
    echo Server configured: http://%EC2_IP%:11434
    echo.
    echo INSTALLATION STEPS:
    echo 1. Open Microsoft Excel
    echo 2. Go to File ^> Options ^> Add-ins
    echo 3. At bottom: Manage: Excel Add-ins ^> Go
    echo 4. Click "Browse" and select ExcelOllamaAIPlugin.xlam
    echo 5. Check the box next to "ExcelOllamaAIPlugin"
    echo 6. Click OK
    echo.
    echo USAGE:
    echo 1. Select your data in Excel (including headers^)
    echo 2. Press Alt+F8 to open Macro dialog
    echo 3. Choose from available functions:
    echo    - AnalyzeSelectedData
    echo    - AskQuestionAboutData
    echo    - GenerateComprehensiveReport
    echo    - ConfigureOllamaServer
    echo    - TestOllamaConnection
    echo.
    echo The add-in will be available every time you open Excel!
    echo.
) else (
    echo.
    echo ❌ ERROR: Add-in creation failed
    echo.
    echo MANUAL STEPS:
    echo 1. Open Excel
    echo 2. Press Alt+F11 (VBA Editor^)
    echo 3. Insert ^> Module
    echo 4. Copy contents of ExcelOllamaAIPlugin.vba
    echo 5. Save as Excel Add-in (.xlam^)
    echo.
)

REM Cleanup temporary files
del build-addin.vbs 2>nul

echo.
pause