# 📁 How to Create Excel Add-in (.xlam) File

## Method 1: Manual Creation (Recommended)

### Step 1: Create New Excel Workbook
1. **Open Microsoft Excel**
2. **Create a new blank workbook**
3. **Press Alt+F11** to open VBA Editor

### Step 2: Add VBA Code
1. **In VBA Editor:** Right-click on VBAProject → Insert → Module
2. **Copy the entire code** from `create-excel-addin.vba`
3. **Update the server IP** in this line:
   ```vba
   Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
   Private Const WEB_SERVER As String = "http://YOUR_EC2_IP:3000"
   ```
   Replace `YOUR_EC2_IP` with your actual EC2 public IP address

### Step 3: Save as Add-in
1. **Press Ctrl+S** or File → Save As
2. **Choose file type:** Excel Add-in (*.xlam)
3. **File name:** ExcelOllamaAIPlugin.xlam
4. **Location:** Excel will suggest the Add-ins folder (keep this location)
5. **Click Save**

### Step 4: Install the Add-in
1. **Close VBA Editor** (Alt+Q)
2. **In Excel:** File → Options → Add-ins
3. **At bottom:** Manage: Excel Add-ins → Go
4. **Check the box** next to "ExcelOllamaAIPlugin"
5. **Click OK**

### Step 5: Use the Add-in
1. **Select your data** in Excel (including headers)
2. **Press Alt+F8** to open Macro dialog
3. **Choose from available functions:**
   - `AnalyzeSelectedData` - Statistical analysis
   - `AskQuestionAboutData` - Natural language queries
   - `GenerateComprehensiveReport` - Multi-analysis report
   - `AnalyzeTrends` - Trend analysis
   - `DetectPatterns` - Pattern detection
   - `ConfigureOllamaServer` - Change server settings
   - `TestOllamaConnection` - Test connectivity

## Method 2: Automated Creation (Windows Script)

### Step 1: Save the VBScript
1. **Copy this code** and save as `create-addin.vbs`:

```vbscript
' VBScript to automatically create Excel Add-in (.xlam)
Dim xl, wb, vbComp, fso, file
Dim vbaCode, serverIP

' Get EC2 server IP from user
serverIP = InputBox("Enter your EC2 Public IP address:", "Server Configuration")
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

' VBA Code (embedded)
vbaCode = "Option Explicit" & vbCrLf & _
"Private Const OLLAMA_SERVER As String = ""http://" & serverIP & ":11434""" & vbCrLf & _
"Private Const WEB_SERVER As String = ""http://" & serverIP & ":3000""" & vbCrLf & _
"Private Const DEFAULT_MODEL As String = ""llama2:latest""" & vbCrLf & vbCrLf & _
"' [REST OF VBA CODE WOULD GO HERE - truncated for brevity]"

' Add VBA module
Set vbComp = wb.VBProject.VBComponents.Add(1)
vbComp.CodeModule.AddFromString vbaCode

' Save as Excel Add-in
wb.SaveAs xl.Application.DefaultFilePath & "\ExcelOllamaAIPlugin.xlam", 18

' Install the add-in
xl.AddIns.Add(xl.Application.DefaultFilePath & "\ExcelOllamaAIPlugin.xlam").Installed = True

MsgBox "Excel-Ollama AI Plugin created and installed successfully!"

' Cleanup
xl.DisplayAlerts = True
xl.Quit
```

### Step 2: Run the Script
1. **Double-click** `create-addin.vbs`
2. **Enter your EC2 IP** when prompted
3. **Wait for completion** message

## Method 3: Pre-built Add-in Download

### Step 1: Build on EC2
```bash
# On your EC2 instance, create the add-in builder
cat > build-excel-addin.sh << 'EOF'
#!/bin/bash
EC2_IP=$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)

# Create VBA code with actual IP
sed "s/YOUR_EC2_IP/$EC2_IP/g" create-excel-addin.vba > ExcelOllamaAIPlugin.vba

echo "Excel Add-in VBA code ready!"
echo "Download ExcelOllamaAIPlugin.vba to your Windows machine"
echo "Then follow Manual Creation steps 1-5"
echo ""
echo "Your server IP: $EC2_IP"
echo "Ollama URL: http://$EC2_IP:11434"
EOF

chmod +x build-excel-addin.sh
./build-excel-addin.sh
```

### Step 2: Download and Create
1. **Download** `ExcelOllamaAIPlugin.vba` from EC2
2. **Follow Manual Creation** steps 1-5 above

## 🎯 Quick Test

After installation:

1. **Open Excel**
2. **Create sample data:**
   ```
   Date        Sales   Product   Region
   2024-01-01  1000    A         North
   2024-01-02  1200    B         South
   2024-01-03  800     A         East
   ```
3. **Select the data** (including headers)
4. **Press Alt+F8**
5. **Run** `AnalyzeSelectedData`
6. **Check** the new "AI_Analysis_Results" sheet

## 🔧 Configuration

### First Time Setup:
1. **Press Alt+F8**
2. **Run** `ConfigureOllamaServer`
3. **Enter your EC2 details:**
   - Server URL: `http://YOUR_EC2_IP:11434`
   - Model: `llama2:latest`
4. **Test connection** with `TestOllamaConnection`

## 🚨 Troubleshooting

### "Macro not found"
- Ensure add-in is installed: File → Options → Add-ins
- Check that "ExcelOllamaAIPlugin" is checked

### "Connection failed"
- Run `TestOllamaConnection` to diagnose
- Verify EC2 instance is running
- Check Security Group allows port 11434

### "VBA errors"
- Enable macros: File → Options → Trust Center → Macro Settings
- Choose "Enable all macros" (for development)

## 📋 Available Functions

| Function | Description |
|----------|-------------|
| `AnalyzeSelectedData` | Statistical analysis of selected data |
| `AskQuestionAboutData` | Ask questions in natural language |
| `GenerateComprehensiveReport` | Multi-perspective analysis report |
| `AnalyzeTrends` | Trend and time series analysis |
| `DetectPatterns` | Pattern and anomaly detection |
| `ConfigureOllamaServer` | Change server and model settings |
| `TestOllamaConnection` | Test server connectivity |
| `ShowHelp` | Display help information |

## 🎉 Success!

Once installed, you'll have:
- ✅ **Native Excel Add-in** (.xlam format)
- ✅ **No Python required** on Windows
- ✅ **Direct API calls** to your EC2 Ollama server
- ✅ **Multiple analysis types** available
- ✅ **Natural language queries** supported
- ✅ **Results in Excel sheets** automatically

The add-in will be available every time you open Excel, and you can use it with any workbook!