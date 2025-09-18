# 📁 How to Create Excel Add-in (.xlam) File

Follow these steps to create a proper Excel Add-in (.xlam) file that you can install in Excel.

## 🎯 Step-by-Step Instructions

### Step 1: Create New Excel Workbook

1. **Open Microsoft Excel**
2. **Create a new blank workbook**
3. **Save it immediately** as "ExcelOllamaAIPlugin.xlsm" (Excel Macro-Enabled Workbook)

### Step 2: Open VBA Editor

1. **Press Alt+F11** to open the VBA Editor
2. **Or go to Developer tab → Visual Basic** (if Developer tab is visible)
3. **If Developer tab is not visible:**
   - File → Options → Customize Ribbon
   - Check "Developer" in the right panel
   - Click OK

### Step 3: Add VBA Code

1. **In VBA Editor, right-click on "VBAProject (ExcelOllamaAIPlugin.xlsm)"**
2. **Select Insert → Module**
3. **Copy all the code from `create-excel-addin.vba`** and paste it into the module
4. **IMPORTANT: Update the server URLs at the top of the code:**
   ```vba
   Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
   Private Const WEB_SERVER As String = "http://YOUR_EC2_IP:3000"
   ```
   Replace `YOUR_EC2_IP` with your actual EC2 public IP address.

### Step 4: Save as Excel Add-in

1. **Press Ctrl+S or File → Save**
2. **In the Save dialog:**
   - **File name:** ExcelOllamaAIPlugin
   - **Save as type:** Excel Add-in (*.xlam)
   - **Location:** Choose a location you'll remember (e.g., Desktop)
3. **Click Save**

### Step 5: Install the Add-in

1. **Close the VBA Editor**
2. **In Excel, go to File → Options → Add-ins**
3. **At the bottom, select "Excel Add-ins" and click "Go..."**
4. **Click "Browse..."**
5. **Navigate to where you saved the .xlam file**
6. **Select "ExcelOllamaAIPlugin.xlam" and click OK**
7. **Make sure the checkbox next to "ExcelOllamaAIPlugin" is checked**
8. **Click OK**

### Step 6: Verify Installation

1. **The add-in should now be loaded**
2. **You should see a welcome message when Excel starts**
3. **To use the functions:**
   - **Press Alt+F8** to see available macros
   - **Or go to Developer → Macros**

## 🚀 How to Use the Add-in

### Basic Usage:

1. **Select your data range** (including headers)
2. **Press Alt+F8** to open Macros dialog
3. **Choose one of these functions:**
   - `AnalyzeSelectedData` - Statistical analysis
   - `AskQuestionAboutData` - Natural language queries
   - `GenerateComprehensiveReport` - Multi-perspective analysis
   - `AnalyzeTrends` - Trend analysis
   - `DetectPatterns` - Pattern detection
4. **Click "Run"**
5. **Results appear in a new sheet**

### Configuration:

1. **Press Alt+F8**
2. **Run `ConfigureOllamaServer`** to set your EC2 server URL
3. **Run `TestOllamaConnection`** to verify connectivity

## 🔧 Alternative: Quick Installation Script

If you prefer, I can create a script that builds the .xlam file automatically:

<function_calls>
<invoke name="fsWrite">
<parameter name="path">build-xlam.vbsipt to automatically create Excel Add-in (.xlam)
' Run this script to automatically build the Excel add-in

Dim xl, wb, vbComp, fso, file
Dim vbaCode, serverIP

' Get EC2 server IP from user
serverIP = InputBox("Enter your EC2 Public IP address:", "Server Configuration", "YOUR_EC2_IP")
If serverIP = "" Then
    MsgBox "No IP address provided. Exiting."
    WScript.Quit
End If

' Create Excel application
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False

' Create new workbook
Set wb = xl.Workbooks.Add

' Read VBA code from file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("create-excel-addin.vba", 1)
vbaCode = file.ReadAll
file.Close

' Replace placeholder with actual IP
vbaCode = Replace(vbaCode, "YOUR_EC2_IP", serverIP)

' Add VBA module
Set vbComp = wb.VBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
vbComp.CodeModule.AddFromString vbaCode

' Save as Excel Add-in
wb.SaveAs "ExcelOllamaAIPlugin.xlam", 18 ' 18 = xlAddIn

' Close workbook
wb.Close False

' Install the add-in
xl.AddIns.Add("ExcelOllamaAIPlugin.xlam").Installed = True

' Show completion message
MsgBox "Excel-Ollama AI Plugin created and installed successfully!" & vbCrLf & vbCrLf & _
       "Server: http://" & serverIP & ":11434" & vbCrLf & _
       "Add-in: ExcelOllamaAIPlugin.xlam" & vbCrLf & vbCrLf & _
       "Press Alt+F8 in Excel to use the functions."

' Cleanup
xl.DisplayAlerts = True
xl.Quit
Set xl = Nothing