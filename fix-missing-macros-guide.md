# Fix Missing Macros After Excel Restart

## Problem
- Plugin shows "loaded" message when Excel opens
- Alt+F8 shows empty macro list
- Functions not available

## Solution Steps

### Step 1: Check Add-in Installation
1. **Open Excel**
2. **Go to File > Options > Add-ins**
3. **At bottom, select "Excel Add-ins" from dropdown**
4. **Click "Go..." button**
5. **Look for "OllamaAI_Enterprise" in the list**
6. **If not there, click "Browse..." and find your .xlam file**
7. **Check the box next to the add-in name**
8. **Click OK**

### Step 2: Check Macro Security Settings
1. **Go to File > Options > Trust Center**
2. **Click "Trust Center Settings..." button**
3. **Click "Macro Settings" on left**
4. **Select "Enable all macros" (temporarily for testing)**
5. **Click OK and restart Excel**

### Step 3: Verify Functions are Public
The VBA functions must be marked as `Public` not `Private`:
```vb
Public Sub ConfigureOllamaServer()  ' ✅ Correct
Private Sub ConfigureOllamaServer() ' ❌ Wrong - won't show in Alt+F8
```

### Step 4: Alternative Access Methods
If Alt+F8 doesn't show functions, try:
1. **Press Alt+F11** (VBA Editor)
2. **Look for your add-in module in Project Explorer**
3. **Double-click the module to see the code**
4. **Place cursor in a function and press F5 to run it**

### Step 5: Manual Function Execution
You can also run functions directly:
1. **Press Ctrl+G** (Immediate Window in VBA Editor)
2. **Type function name and press Enter:**
   ```
   ConfigureOllamaServer
   TestConnection
   AskQuestionAboutDataEnterprise
   ```

## Quick Test
1. **Open Excel**
2. **Press Alt+F11** (VBA Editor)
3. **Press Ctrl+G** (Immediate Window)
4. **Type: `ConfigureOllamaServer` and press Enter**
5. **If it runs, the add-in is working**