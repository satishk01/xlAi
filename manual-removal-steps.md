# Manual Plugin Removal Steps

## Complete Removal Process

### Step 1: Remove from Excel Add-ins List
1. **Open Excel**
2. **Go to File > Options > Add-ins**
3. **At bottom, select "Excel Add-ins" from dropdown**
4. **Click "Go..." button**
5. **Look for "OllamaAI_Enterprise" or similar**
6. **Uncheck the box next to it**
7. **If you see a "missing file" error, select it and click "Remove"**
8. **Click OK**

### Step 2: Delete Physical Files
Navigate to these locations and delete any .xlam files:

**Excel Add-ins Directory:**
```
%APPDATA%\Microsoft\AddIns\
```
Look for files like:
- OllamaAI_Enterprise.xlam
- enterprise-excel-addin.xlam
- bulletproof-excel-addin.xlam

**Your Working Directory:**
Delete any .xlam files in the folder where you ran the deployment script.

### Step 3: Remove VBA Code (if manually added)
If you manually copied VBA code into a workbook:
1. **Open the workbook**
2. **Press Alt+F11** (VBA Editor)
3. **In Project Explorer, find modules with Ollama code**
4. **Right-click the module > Remove**
5. **Choose "No" when asked to export**
6. **Save the workbook**

### Step 4: Clear Temporary Files
Delete these files from your working directory:
- create_enterprise_addin.ps1
- create_enterprise_addin_fixed.ps1
- temp_addin.xlsm
- ENTERPRISE_SETUP_INSTRUCTIONS.txt

### Step 5: Restart Excel
1. **Close Excel completely**
2. **Reopen Excel**
3. **Verify no "plugin loaded" message appears**
4. **Press Alt+F8 - should show no Ollama functions**

## Verification Checklist
- [ ] No "plugin loaded" message on Excel startup
- [ ] Alt+F8 shows no Ollama functions
- [ ] File > Options > Add-ins shows no Ollama add-ins
- [ ] No .xlam files in AddIns directory
- [ ] No VBA modules with Ollama code in workbooks

## If Plugin Still Appears
1. **Check Excel Trust Center settings**
2. **Look for .xlam files in other Office directories**
3. **Restart Windows (last resort)**
4. **Check for multiple Excel installations**