' ============================================================================
' Excel-Ollama AI Plugin - FIXED VERSION for Runtime Error 1004
' Handles Excel object access issues properly
' ============================================================================

Option Explicit

' Configuration - UPDATE THIS WITH YOUR EC2 IP
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const DEFAULT_MODEL As String = "llama2:latest"

' Global variables
Private currentModel As String
Private serverUrl As String

' ============================================================================
' INITIALIZATION
' ============================================================================
Sub Auto_Open()
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    
    MsgBox "🤖 Excel-Ollama AI Plugin (Fixed Version) loaded!" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel, vbInformation, "Ollama AI Plugin"
End Sub

' ============================================================================
' MAIN FUNCTIONS - FIXED FOR RUNTIME ERROR 1004
' ============================================================================

' 1. Analyze selected data - FIXED VERSION
Public Sub AnalyzeSelectedData()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    Dim rowCount As Long, colCount As Long
    
    ' Safely get selected range
    Set selectedRange = Application.Selection
    
    ' Validate selection
    If selectedRange Is Nothing Then
        MsgBox "❌ No range selected. Please select your data first.", vbExclamation, "Selection Required"
        Exit Sub
    End If
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    If rowCount < 2 Then
        MsgBox "❌ Please select at least 2 rows (including headers)." & vbCrLf & vbCrLf & _
               "Current selection: " & rowCount & " rows", vbExclamation, "Insufficient Data"
        Exit Sub
    End If
    
    If colCount > 20 Then
        MsgBox "⚠️ Large selection detected (" & colCount & " columns)." & vbCrLf & _
               "This may take longer to process.", vbInformation, "Large Dataset"
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Analyzing " & rowCount & " rows × " & colCount & " columns..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Safely get data as array
    dataArray = selectedRange.Value2  ' Use Value2 for better performance
    
    ' Perform analysis with error handling
    analysisResult = CallOllamaAPISafe(dataArray, "statistical")
    
    ' Write results to new sheet safely
    Call WriteResultsToSheetSafe(analysisResult, "AI_Analysis_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "✅ Analysis completed!" & vbCrLf & vbCrLf & _
           "Data processed: " & rowCount & " rows × " & colCount & " columns" & vbCrLf & _
           "Results: AI_Analysis_Results sheet", vbInformation, "Analysis Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "❌ Runtime Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "This usually means:" & vbCrLf & _
           "• Invalid range selection" & vbCrLf & _
           "• Worksheet access issue" & vbCrLf & _
           "• Memory limitation" & vbCrLf & vbCrLf & _
           "Try selecting a smaller data range.", vbCritical, "Runtime Error 1004"
End Sub

' 2. Ask question - FIXED VERSION
Public Sub AskQuestionAboutData()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim question As String
    Dim answer As String
    Dim rowCount As Long, colCount As Long
    
    ' Get user question first
    question = InputBox("Ask a simple question about your data:" & vbCrLf & vbCrLf & _
                       "Good examples:" & vbCrLf & _
                       "• How many rows of data?" & vbCrLf & _
                       "• What columns are available?" & vbCrLf & _
                       "• What is the average of column X?" & vbCrLf & _
                       "• Summarize this data", "🤖 Ask Question")
    
    If question = "" Or Len(question) < 3 Then
        MsgBox "❌ Please enter a valid question.", vbExclamation, "Question Required"
        Exit Sub
    End If
    
    ' Safely get selected range
    Set selectedRange = Application.Selection
    
    If selectedRange Is Nothing Then
        MsgBox "❌ No data selected. Please select your data range first.", vbExclamation, "Selection Required"
        Exit Sub
    End If
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    If rowCount < 2 Then
        MsgBox "❌ Please select at least 2 rows (including headers).", vbExclamation, "Insufficient Data"
        Exit Sub
    End If
    
    ' Limit data size to prevent issues
    If rowCount > 50 Then
        If MsgBox("Large dataset detected (" & rowCount & " rows)." & vbCrLf & vbCrLf & _
                  "For better performance, consider using first 50 rows only." & vbCrLf & vbCrLf & _
                  "Continue with full dataset?", vbYesNo + vbQuestion, "Large Dataset") = vbNo Then
            ' Use only first 50 rows
            Set selectedRange = selectedRange.Resize(50, colCount)
            rowCount = 50
        End If
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Processing question: " & Left(question, 30) & "..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Safely get data as array
    dataArray = selectedRange.Value2
    
    ' Ask question with enhanced error handling
    answer = AskOllamaQuestionSafe(dataArray, question)
    
    ' Create result text
    Dim resultText As String
    resultText = "QUESTION: " & question & vbCrLf
    resultText = resultText & String(60, "=") & vbCrLf & vbCrLf
    resultText = resultText & "DATA INFO:" & vbCrLf
    resultText = resultText & "Rows: " & rowCount & ", Columns: " & colCount & vbCrLf & vbCrLf
    resultText = resultText & "ANSWER:" & vbCrLf
    resultText = resultText & String(30, "-") & vbCrLf
    resultText = resultText & answer
    
    ' Write results safely
    Call WriteResultsToSheetSafe(resultText, "AI_Question_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "✅ Question answered!" & vbCrLf & vbCrLf & _
           "Question: " & Left(question, 50) & "..." & vbCrLf & _
           "Results: AI_Question_Results sheet", vbInformation, "Question Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "❌ Runtime Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Common causes:" & vbCrLf & _
           "• Data range too large (try smaller selection)" & vbCrLf & _
           "• Invalid characters in data" & vbCrLf & _
           "• Worksheet protection enabled" & vbCrLf & _
           "• Memory limitation" & vbCrLf & vbCrLf & _
           "Try selecting a smaller, simpler data range.", vbCritical, "Runtime Error 1004"
End Sub

' 3. Simple test function
Public Sub TestWithSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim testRange As Range
    Dim testData As Variant
    
    ' Create sample data in current sheet
    Set ws = ActiveSheet
    
    ' Clear a small area first
    ws.Range("A1:C4").Clear
    
    ' Create simple test data
    ws.Range("A1").Value = "Name"
    ws.Range("B1").Value = "Age"
    ws.Range("C1").Value = "Score"
    ws.Range("A2").Value = "John"
    ws.Range("B2").Value = 25
    ws.Range("C2").Value = 85
    ws.Range("A3").Value = "Jane"
    ws.Range("B3").Value = 30
    ws.Range("C3").Value = 92
    ws.Range("A4").Value = "Bob"
    ws.Range("B4").Value = 35
    ws.Range("C4").Value = 78
    
    ' Select the test data
    Set testRange = ws.Range("A1:C4")
    testRange.Select
    
    ' Get data safely
    testData = testRange.Value2
    
    ' Test the API call
    Application.StatusBar = "🧪 Testing with sample data..."
    
    Dim result As String
    result = AskOllamaQuestionSafe(testData, "What is the average age and score?")
    
    ' Show result in message box
    Application.StatusBar = False
    
    MsgBox "🧪 Test Result:" & vbCrLf & vbCrLf & result, vbInformation, "Sample Data Test"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "❌ Test Error " & Err.Number & ": " & Err.Description, vbCritical, "Test Failed"
End Sub

' ============================================================================
' SAFE API FUNCTIONS - FIXED FOR RUNTIME ERROR 1004
' ============================================================================

' Safe API call with comprehensive error handling
Private Function CallOllamaAPISafe(dataArray As Variant, analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim prompt As String
    
    ' Validate input data
    If IsEmpty(dataArray) Then
        CallOllamaAPISafe = "❌ Error: No data provided"
        Exit Function
    End If
    
    ' Check array bounds safely
    Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
    minRow = LBound(dataArray, 1)
    maxRow = UBound(dataArray, 1)
    minCol = LBound(dataArray, 2)
    maxCol = UBound(dataArray, 2)
    
    If maxRow - minRow < 1 Then
        CallOllamaAPISafe = "❌ Error: Need at least 2 rows of data"
        Exit Function
    End If
    
    ' Build prompt safely
    prompt = BuildPromptSafe(dataArray, analysisType)
    
    If Len(prompt) = 0 Then
        CallOllamaAPISafe = "❌ Error: Could not build analysis prompt"
        Exit Function
    End If
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build request body with safe JSON
    requestBody = BuildSafeJSONRequest(currentModel, prompt)
    
    ' Make API call
    url = serverUrl & "/api/generate"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPISafe = ExtractResponseSafe(response)
    Else
        CallOllamaAPISafe = "❌ HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & vbCrLf & _
                           "Server: " & serverUrl & vbCrLf & _
                           "Please check server configuration."
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPISafe = "❌ API Error: " & Err.Description & " (Error " & Err.Number & ")" & vbCrLf & vbCrLf & _
                       "Server: " & serverUrl & vbCrLf & _
                       "Model: " & currentModel
End Function

' Safe question asking
Private Function AskOllamaQuestionSafe(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If IsEmpty(dataArray) Then
        AskOllamaQuestionSafe = "❌ Error: No data provided"
        Exit Function
    End If
    
    If Len(Trim(question)) = 0 Then
        AskOllamaQuestionSafe = "❌ Error: No question provided"
        Exit Function
    End If
    
    ' Build simple question prompt
    Dim prompt As String
    prompt = BuildQuestionPromptSafe(dataArray, question)
    
    ' Call API safely
    AskOllamaQuestionSafe = CallOllamaWithPromptSafe(prompt)
    
    Exit Function
    
ErrorHandler:
    AskOllamaQuestionSafe = "❌ Question Error: " & Err.Description & " (Error " & Err.Number & ")"
End Function

' ============================================================================
' SAFE HELPER FUNCTIONS
' ============================================================================

' Build prompt safely without causing array errors
Private Function BuildPromptSafe(dataArray As Variant, analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    ' Get array dimensions safely
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers safely
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        
        ' Safe string conversion
        Dim headerValue As String
        headerValue = CStr(dataArray(LBound(dataArray, 1), j))
        If Len(headerValue) > 50 Then headerValue = Left(headerValue, 50) & "..."
        headers = headers & headerValue
    Next j
    
    ' Extract limited sample data (max 3 rows, 5 columns)
    Dim maxSampleRows As Long, maxSampleCols As Long
    maxSampleRows = Application.Min(3, rowCount - 1)  ' Exclude header
    maxSampleCols = Application.Min(5, colCount)
    
    For i = LBound(dataArray, 1) + 1 To LBound(dataArray, 1) + maxSampleRows
        sampleData = sampleData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To LBound(dataArray, 2) + maxSampleCols - 1
            If j > LBound(dataArray, 2) Then sampleData = sampleData & ", "
            
            Dim cellValue As String
            cellValue = CStr(dataArray(i, j))
            If Len(cellValue) > 20 Then cellValue = Left(cellValue, 20) & "..."
            sampleData = sampleData & cellValue
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build concise prompt
    prompt = "Dataset: " & (rowCount - 1) & " rows, " & colCount & " columns" & vbCrLf
    prompt = prompt & "Headers: " & headers & vbCrLf
    prompt = prompt & "Sample:" & vbCrLf & sampleData & vbCrLf
    
    Select Case analysisType
        Case "statistical"
            prompt = prompt & "Provide statistical summary: averages, patterns, insights."
        Case "trends"
            prompt = prompt & "Analyze trends and patterns in the data."
        Case Else
            prompt = prompt & "Analyze this data and provide key insights."
    End Select
    
    BuildPromptSafe = prompt
    Exit Function
    
ErrorHandler:
    BuildPromptSafe = "Error building prompt: " & Err.Description
End Function

' Build question prompt safely
Private Function BuildQuestionPromptSafe(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers safely (limit to first 10 columns)
    Dim maxCols As Long
    maxCols = Application.Min(10, colCount)
    
    For j = LBound(dataArray, 2) To LBound(dataArray, 2) + maxCols - 1
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    ' Build simple prompt
    prompt = "Data: " & (rowCount - 1) & " rows with columns: " & headers & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf
    prompt = prompt & "Answer based on the data structure described above."
    
    BuildQuestionPromptSafe = prompt
    Exit Function
    
ErrorHandler:
    BuildQuestionPromptSafe = "Error: " & Err.Description
End Function

' Safe API call with prompt
Private Function CallOllamaWithPromptSafe(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' Validate prompt
    If Len(prompt) = 0 Then
        CallOllamaWithPromptSafe = "❌ Error: Empty prompt"
        Exit Function
    End If
    
    ' Limit prompt size to prevent issues
    If Len(prompt) > 2000 Then
        prompt = Left(prompt, 2000) & "... [truncated for processing]"
    End If
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build safe JSON request
    requestBody = BuildSafeJSONRequest(currentModel, prompt)
    url = serverUrl & "/api/generate"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaWithPromptSafe = ExtractResponseSafe(response)
    Else
        CallOllamaWithPromptSafe = "❌ HTTP " & http.Status & ": " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaWithPromptSafe = "❌ API Error: " & Err.Description
End Function

' Build JSON request safely
Private Function BuildSafeJSONRequest(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    ' Safe JSON escaping - handle only essential characters
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, " ")  ' Replace line breaks with spaces
    escapedPrompt = Replace(escapedPrompt, vbCr, " ")
    escapedPrompt = Replace(escapedPrompt, vbLf, " ")
    escapedPrompt = Replace(escapedPrompt, vbTab, " ")
    
    ' Remove multiple spaces
    Do While InStr(escapedPrompt, "  ") > 0
        escapedPrompt = Replace(escapedPrompt, "  ", " ")
    Loop
    
    BuildSafeJSONRequest = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false}"
End Function

' Extract response safely
Private Function ExtractResponseSafe(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim startPos As Long
    Dim endPos As Long
    Dim result As String
    
    ' Simple but robust JSON parsing
    startPos = InStr(jsonText, """response"":""")
    
    If startPos > 0 Then
        startPos = startPos + 12  ' Skip "response":"
        
        ' Find closing quote (handle escaped quotes)
        endPos = startPos
        Do While endPos <= Len(jsonText)
            If Mid(jsonText, endPos, 1) = """" And Mid(jsonText, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        
        If endPos > startPos And endPos <= Len(jsonText) Then
            result = Mid(jsonText, startPos, endPos - startPos)
            ' Basic unescaping
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            ExtractResponseSafe = result
        Else
            ExtractResponseSafe = "❌ Could not parse response (invalid JSON structure)"
        End If
    Else
        ExtractResponseSafe = "❌ No response field found" & vbCrLf & vbCrLf & _
                             "Raw response (first 300 chars):" & vbCrLf & Left(jsonText, 300)
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractResponseSafe = "❌ JSON parsing error: " & Err.Description
End Function

' Write results to sheet safely - FIXED FOR RUNTIME ERROR 1004
Private Sub WriteResultsToSheetSafe(results As String, sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim resultLines As Variant
    Dim i As Long
    Dim safeSheetName As String
    
    ' Create safe sheet name (Excel has restrictions)
    safeSheetName = sheetName
    safeSheetName = Replace(safeSheetName, "[", "_")
    safeSheetName = Replace(safeSheetName, "]", "_")
    safeSheetName = Replace(safeSheetName, "*", "_")
    safeSheetName = Replace(safeSheetName, "?", "_")
    safeSheetName = Replace(safeSheetName, "/", "_")
    safeSheetName = Replace(safeSheetName, "\", "_")
    
    ' Limit sheet name length
    If Len(safeSheetName) > 31 Then
        safeSheetName = Left(safeSheetName, 31)
    End If
    
    ' Delete existing sheet safely
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(safeSheetName).Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Create new sheet
    Set ws = Worksheets.Add
    ws.Name = safeSheetName
    
    ' Split results into lines
    resultLines = Split(results, vbCrLf)
    
    ' Write results line by line (safer than bulk write)
    For i = 0 To UBound(resultLines)
        If i < 1048576 Then  ' Excel row limit
            ws.Cells(i + 1, 1).Value = resultLines(i)
        End If
    Next i
    
    ' Format safely
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 80  ' Fixed width to avoid AutoFit issues
    End With
    
    ' Activate sheet safely
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    
    ' If sheet creation fails, try to write to active sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "Results (sheet creation failed):"
    ActiveSheet.Range("A2").Value = results
    On Error GoTo 0
    
    MsgBox "⚠️ Could not create new sheet. Results written to current sheet.", vbExclamation, "Sheet Creation Issue"
End Sub

' ============================================================================
' CONFIGURATION AND TEST FUNCTIONS
' ============================================================================

' Configure server
Public Sub ConfigureOllamaServer()
    Dim newServer As String
    Dim newModel As String
    
    newServer = InputBox("Enter Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Format: http://your-ec2-ip:11434", _
                        "🔧 Server Configuration", serverUrl)
    
    If newServer <> "" Then
        serverUrl = newServer
        
        newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                           "Examples:" & vbCrLf & _
                           "• llama2:latest" & vbCrLf & _
                           "• mistral:latest", _
                           "🔧 Model Configuration", currentModel)
        
        If newModel <> "" Then
            currentModel = newModel
        End If
        
        MsgBox "✅ Configuration updated!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel, vbInformation, "Configuration Updated"
    End If
End Sub

' Test connection
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    Application.StatusBar = "🔍 Testing connection..."
    
    http.Open "GET", url, False
    http.send
    
    Application.StatusBar = False
    
    If http.Status = 200 Then
        MsgBox "✅ Connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Response: " & Left(http.responseText, 100) & "...", vbInformation, "Connection Test"
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & _
               "Error: " & http.statusText, vbCritical, "Connection Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Connection Error"
End Sub

' Show help
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "🤖 EXCEL-OLLAMA AI PLUGIN (FIXED VERSION)" & vbCrLf & vbCrLf
    helpText = helpText & "🔧 SETUP FUNCTIONS:" & vbCrLf
    helpText = helpText & "• ConfigureOllamaServer - Set your EC2 server URL" & vbCrLf
    helpText = helpText & "• TestConnection - Test server connectivity" & vbCrLf
    helpText = helpText & "• TestWithSampleData - Test with built-in sample data" & vbCrLf & vbCrLf
    helpText = helpText & "📊 ANALYSIS FUNCTIONS:" & vbCrLf
    helpText = helpText & "• AnalyzeSelectedData - Statistical analysis" & vbCrLf
    helpText = helpText & "• AskQuestionAboutData - Natural language questions" & vbCrLf & vbCrLf
    helpText = helpText & "⚙️ CURRENT SETTINGS:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel & vbCrLf & vbCrLf
    helpText = helpText & "🚀 QUICK START:" & vbCrLf
    helpText = helpText & "1. Run ConfigureOllamaServer (set your EC2 IP)" & vbCrLf
    helpText = helpText & "2. Run TestConnection (verify it works)" & vbCrLf
    helpText = helpText & "3. Run TestWithSampleData (test functionality)" & vbCrLf
    helpText = helpText & "4. Select your data and run AnalyzeSelectedData" & vbCrLf & vbCrLf
    helpText = helpText & "💡 TIP: Start with TestWithSampleData to verify everything works!"
    
    MsgBox helpText, vbInformation, "Plugin Help"
End Sub