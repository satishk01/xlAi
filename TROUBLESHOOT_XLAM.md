# 🔧 Troubleshooting Excel Add-in Functions

## Issue: Functions not showing in Alt+F8

### Step 1: Check Add-in Status

1. **Open Excel**
2. **Go to:** File → Options → Add-ins
3. **At bottom:** Manage: Excel Add-ins → Go
4. **Verify:** "ExcelOllamaAIPlugin" is checked ✅
5. **If not listed:** Click Browse and find your .xlam file

### Step 2: Check Macro Security Settings

1. **Go to:** File → Options → Trust Center → Trust Center Settings
2. **Click:** Macro Settings
3. **Select:** "Enable all macros" (temporarily for testing)
4. **Click OK** and restart Excel

### Step 3: Verify VBA Code is Present

1. **Press Alt+F11** (VBA Editor)
2. **Look for:** VBAProject (ExcelOllamaAIPlugin.xlam)
3. **Expand it** and check if Module1 exists
4. **Double-click Module1** to see the code

### Step 4: Check Function Visibility

The functions must be marked as `Public`. In VBA Editor, verify each function starts with:
```vba
Public Sub FunctionName()
```
NOT:
```vba
Private Sub FunctionName()
```

## Quick Fix Solutions:

### Solution 1: Re-create with Corrected Code

1. **Delete current add-in:** File → Options → Add-ins → Uncheck it
2. **Use the corrected VBA code below**
3. **Re-save as .xlam**
4. **Re-install the add-in**

### Solution 2: Direct Access Method

Instead of Alt+F8, try:
1. **Developer tab** → Macros
2. **Or:** View → Macros → View Macros
3. **In "Macros in" dropdown:** Select "ExcelOllamaAIPlugin.xlam"

### Solution 3: Enable Developer Tab

1. **File → Options → Customize Ribbon**
2. **Check "Developer"** in right panel
3. **Click OK**
4. **Use Developer → Macros** instead of Alt+F8