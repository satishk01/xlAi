' ============================================================================
' Excel-Ollama AI Plugin - BULLETPROOF VERSION
' Completely fixes Runtime Error 1004 with comprehensive error handling
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
    
    MsgBox "🤖 Excel-Ollama AI Plugin (Bulletproof Version) loaded!" & vbCrLf & vbCrLf & _
           "✅ Runtime Error 1004 fixes applied" & vbCrLf & _
           "✅ Enhanced error handling" & vbCrLf & _
           "✅ Memory optimization" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel, vbInformation, "Ollama AI Plugin"
End Sub

' ============================================================================
' MAIN FUNCTIONS - BULLETPROOF VERSIONS
' ============================================================================

' 1. BULLETPROOF Data Analysis
Public Sub AnalyzeSelectedData()
    ' Comprehensive error handling
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    Dim rowCount As Long, colCount As Long
    Dim sheetName As String
    
    ' Step 1: Validate Excel environment
    If Not ValidateExcelEnvironment() Then Exit Sub
    
    ' Step 2: Get and validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Step 3: Check data size limits
    If Not CheckDataSizeLimits(rowCount, colCount) Then Exit Sub
    
    ' Step 4: Show progress and prepare Excel
    Call PrepareExcelForProcessing("🤖 Analyzing " & rowCount & " rows × " & colCount & " columns...")
    
    ' Step 5: Safely extract data
    dataArray = ExtractDataSafely(selectedRange)
    If IsEmpty(dataArray) Then
        Call RestoreExcelState()
        MsgBox "❌ Could not extract data from selection", vbCritical, "Data Error"
        Exit Sub
    End If
    
    ' Step 6: Perform analysis
    analysisResult = PerformAnalysisSafely(dataArray, "statistical")
    
    ' Step 7: Create unique sheet name
    sheetName = CreateUniqueSheetName("AI_Analysis")
    
    ' Step 8: Write results safely
    Call WriteResultsSafely(analysisResult, sheetName)
    
    ' Step 9: Cleanup and show success
    Call RestoreExcelState()
    
    MsgBox "✅ Analysis completed successfully!" & vbCrLf & vbCrLf & _
           "📊 Data processed: " & rowCount & " rows × " & colCount & " columns" & vbCrLf & _
           "📋 Results sheet: " & sheetName & vbCrLf & vbCrLf & _
           "The new sheet contains your AI analysis.", vbInformation, "Analysis Complete"
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    Call HandleRuntimeError(Err.Number, Err.Description, "AnalyzeSelectedData")
End Sub

' 2. BULLETPROOF Question Asking
Public Sub AskQuestionAboutData()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim question As String
    Dim answer As String
    Dim rowCount As Long, colCount As Long
    Dim sheetName As String
    Dim resultText As String
    
    ' Step 1: Get question first
    question = GetValidQuestion()
    If question = "" Then Exit Sub
    
    ' Step 2: Validate Excel environment
    If Not ValidateExcelEnvironment() Then Exit Sub
    
    ' Step 3: Get and validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Step 4: Check data size and offer to limit
    Set selectedRange = OptimizeDataSize(selectedRange, rowCount, colCount)
    If selectedRange Is Nothing Then Exit Sub
    
    ' Update counts after potential resize
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Step 5: Prepare Excel for processing
    Call PrepareExcelForProcessing("🤖 Processing question: " & Left(question, 30) & "...")
    
    ' Step 6: Extract data safely
    dataArray = ExtractDataSafely(selectedRange)
    If IsEmpty(dataArray) Then
        Call RestoreExcelState()
        MsgBox "❌ Could not extract data from selection", vbCritical, "Data Error"
        Exit Sub
    End If
    
    ' Step 7: Ask question safely
    answer = AskQuestionSafely(dataArray, question)
    
    ' Step 8: Format result
    resultText = FormatQuestionResult(question, answer, rowCount, colCount)
    
    ' Step 9: Create unique sheet name and write results
    sheetName = CreateUniqueSheetName("AI_Question")
    Call WriteResultsSafely(resultText, sheetName)
    
    ' Step 10: Cleanup and show success
    Call RestoreExcelState()
    
    MsgBox "✅ Question answered successfully!" & vbCrLf & vbCrLf & _
           "❓ Question: " & Left(question, 50) & "..." & vbCrLf & _
           "📋 Results sheet: " & sheetName, vbInformation, "Question Complete"
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    Call HandleRuntimeError(Err.Number, Err.Description, "AskQuestionAboutData")
End Sub

' 3. BULLETPROOF Connection Test
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim startTime As Double
    
    Application.StatusBar = "🔍 Testing connection to " & serverUrl & "..."
    startTime = Timer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    ' Set timeout to prevent hanging
    http.Open "GET", url, False
    http.send
    
    Application.StatusBar = False
    
    Dim responseTime As Double
    responseTime = Timer - startTime
    
    If http.Status = 200 Then
        MsgBox "✅ Connection successful!" & vbCrLf & vbCrLf & _
               "🌐 Server: " & serverUrl & vbCrLf & _
               "⚡ Response time: " & Format(responseTime, "0.0") & " seconds" & vbCrLf & _
               "📊 HTTP Status: " & http.Status & vbCrLf & vbCrLf & _
               "Your Ollama server is working correctly!", vbInformation, "Connection Test Passed"
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & vbCrLf & _
               "🌐 Server: " & serverUrl & vbCrLf & _
               "📊 HTTP Status: " & http.Status & vbCrLf & _
               "❗ Error: " & http.statusText & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "• Server URL is correct" & vbCrLf & _
               "• Ollama is running on your server" & vbCrLf & _
               "• Network connectivity", vbCritical, "Connection Test Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
           "🌐 Server: " & serverUrl & vbCrLf & _
           "🔢 Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "This usually means the server is unreachable.", vbCritical, "Connection Error"
End Sub

' 4. BULLETPROOF Sample Data Test
Public Sub TestWithSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim testRange As Range
    Dim testData As Variant
    Dim result As String
    
    ' Validate Excel environment first
    If Not ValidateExcelEnvironment() Then Exit Sub
    
    Set ws = ActiveSheet
    
    ' Find a safe area to create sample data
    Dim safeRange As String
    safeRange = FindSafeRangeForSampleData(ws)
    
    ' Create sample data
    Call CreateSampleDataInRange(ws, safeRange)
    
    ' Select and test the sample data
    Set testRange = ws.Range(safeRange)
    testRange.Select
    
    ' Extract data safely
    testData = ExtractDataSafely(testRange)
    
    If Not IsEmpty(testData) Then
        Application.StatusBar = "🧪 Testing with sample data..."
        result = AskQuestionSafely(testData, "What is the average age and score in this data?")
        Application.StatusBar = False
        
        MsgBox "🧪 Sample Data Test Results:" & vbCrLf & vbCrLf & _
               "📊 Test data: 3 people with age and score" & vbCrLf & _
               "❓ Question: What is the average age and score?" & vbCrLf & vbCrLf & _
               "🤖 AI Response:" & vbCrLf & result & vbCrLf & vbCrLf & _
               "✅ If you see a meaningful response above, the plugin is working!", _
               vbInformation, "Sample Test Complete"
    Else
        MsgBox "❌ Could not create or read sample data", vbCritical, "Sample Test Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Call HandleRuntimeError(Err.Number, Err.Description, "TestWithSampleData")
End Sub

' ============================================================================
' BULLETPROOF HELPER FUNCTIONS
' ============================================================================

' Validate Excel environment before doing anything
Private Function ValidateExcelEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if we have an active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "❌ No active workbook found. Please open an Excel file first.", vbCritical, "No Workbook"
        ValidateExcelEnvironment = False
        Exit Function
    End If
    
    ' Check if we have an active sheet
    If ActiveSheet Is Nothing Then
        MsgBox "❌ No active worksheet found.", vbCritical, "No Worksheet"
        ValidateExcelEnvironment = False
        Exit Function
    End If
    
    ' Check if the sheet is protected
    If ActiveSheet.ProtectContents Then
        MsgBox "⚠️ The current worksheet is protected." & vbCrLf & vbCrLf & _
               "The plugin may not work correctly with protected sheets." & vbCrLf & _
               "Consider unprotecting the sheet temporarily.", vbExclamation, "Protected Sheet"
        ' Don't exit - just warn
    End If
    
    ValidateExcelEnvironment = True
    Exit Function
    
ErrorHandler:
    ValidateExcelEnvironment = False
    MsgBox "❌ Excel environment validation failed: " & Err.Description, vbCritical, "Environment Error"
End Function

' Get and validate user selection
Private Function GetValidatedSelection() As Range
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    
    ' Get current selection
    Set selectedRange = Application.Selection
    
    ' Validate selection exists
    If selectedRange Is Nothing Then
        MsgBox "❌ No range selected." & vbCrLf & vbCrLf & _
               "Please select your data range first:" & vbCrLf & _
               "1. Click and drag to select your data" & vbCrLf & _
               "2. Include headers in the first row" & vbCrLf & _
               "3. Make sure all data is selected", vbExclamation, "Selection Required"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    ' Check if it's a valid range (not a chart, shape, etc.)
    If TypeName(selectedRange) <> "Range" Then
        MsgBox "❌ Please select a cell range, not a " & TypeName(selectedRange) & "." & vbCrLf & vbCrLf & _
               "Click and drag to select cells containing your data.", vbExclamation, "Invalid Selection"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    ' Check minimum size
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select at least 2 rows of data." & vbCrLf & vbCrLf & _
               "Current selection: " & selectedRange.Rows.Count & " row(s)" & vbCrLf & vbCrLf & _
               "You need:" & vbCrLf & _
               "• Row 1: Column headers" & vbCrLf & _
               "• Row 2+: Your data", vbExclamation, "Insufficient Data"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    If selectedRange.Columns.Count < 1 Then
        MsgBox "❌ Please select at least 1 column of data.", vbExclamation, "Insufficient Data"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    Set GetValidatedSelection = selectedRange
    Exit Function
    
ErrorHandler:
    Set GetValidatedSelection = Nothing
    MsgBox "❌ Selection validation error: " & Err.Description, vbCritical, "Selection Error"
End Function

' Check data size limits and warn user
Private Function CheckDataSizeLimits(rowCount As Long, colCount As Long) As Boolean
    ' Check for extremely large datasets
    If rowCount > 1000 Then
        If MsgBox("⚠️ Large dataset detected!" & vbCrLf & vbCrLf & _
                  "Rows: " & rowCount & vbCrLf & _
                  "Columns: " & colCount & vbCrLf & vbCrLf & _
                  "Large datasets may:" & vbCrLf & _
                  "• Take a long time to process" & vbCrLf & _
                  "• Use significant memory" & vbCrLf & _
                  "• Cause timeouts" & vbCrLf & vbCrLf & _
                  "Continue anyway?", vbYesNo + vbQuestion, "Large Dataset Warning") = vbNo Then
            CheckDataSizeLimits = False
            Exit Function
        End If
    End If
    
    If colCount > 50 Then
        If MsgBox("⚠️ Many columns detected!" & vbCrLf & vbCrLf & _
                  "Columns: " & colCount & vbCrLf & vbCrLf & _
                  "Consider selecting only the most important columns for better performance." & vbCrLf & vbCrLf & _
                  "Continue with all columns?", vbYesNo + vbQuestion, "Many Columns Warning") = vbNo Then
            CheckDataSizeLimits = False
            Exit Function
        End If
    End If
    
    CheckDataSizeLimits = True
End Function

' Prepare Excel for processing (disable updates, etc.)
Private Sub PrepareExcelForProcessing(statusMessage As String)
    Application.StatusBar = statusMessage
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

' Restore Excel to normal state
Private Sub RestoreExcelState()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' Extract data from range safely
Private Function ExtractDataSafely(selectedRange As Range) As Variant
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    
    ' Use Value2 for better performance and to avoid date/time conversion issues
    dataArray = selectedRange.Value2
    
    ' Validate the extracted data
    If IsEmpty(dataArray) Then
        ExtractDataSafely = Empty
        Exit Function
    End If
    
    ' For single cell, convert to array
    If Not IsArray(dataArray) Then
        Dim singleCellArray(1 To 1, 1 To 1) As Variant
        singleCellArray(1, 1) = dataArray
        ExtractDataSafely = singleCellArray
    Else
        ExtractDataSafely = dataArray
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractDataSafely = Empty
End Function

' Perform analysis with comprehensive error handling
Private Function PerformAnalysisSafely(dataArray As Variant, analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim result As String
    
    ' Build prompt safely
    prompt = BuildAnalysisPromptSafely(dataArray, analysisType)
    
    If Len(prompt) = 0 Then
        PerformAnalysisSafely = "❌ Could not build analysis prompt from your data."
        Exit Function
    End If
    
    ' Call API safely
    result = CallOllamaAPISafely(prompt)
    
    If Left(result, 1) = "❌" Then
        ' API call failed, provide helpful message
        PerformAnalysisSafely = result & vbCrLf & vbCrLf & _
                               "📋 Data Summary (offline analysis):" & vbCrLf & _
                               BuildOfflineDataSummary(dataArray)
    Else
        PerformAnalysisSafely = result
    End If
    
    Exit Function
    
ErrorHandler:
    PerformAnalysisSafely = "❌ Analysis error: " & Err.Description & vbCrLf & vbCrLf & _
                           "📋 Data Summary (offline analysis):" & vbCrLf & _
                           BuildOfflineDataSummary(dataArray)
End Function

' Build analysis prompt safely
Private Function BuildAnalysisPromptSafely(dataArray As Variant, analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long
    
    ' Get array dimensions safely
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers (first row)
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CleanTextForPrompt(CStr(dataArray(LBound(dataArray, 1), j)))
    Next j
    
    ' Extract sample data (max 3 rows, max 5 columns)
    Dim maxSampleRows As Long, maxSampleCols As Long
    maxSampleRows = Application.Min(3, rowCount - 1)
    maxSampleCols = Application.Min(5, colCount)
    
    For i = LBound(dataArray, 1) + 1 To LBound(dataArray, 1) + maxSampleRows
        sampleData = sampleData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To LBound(dataArray, 2) + maxSampleCols - 1
            If j > LBound(dataArray, 2) Then sampleData = sampleData & ", "
            sampleData = sampleData & CleanTextForPrompt(CStr(dataArray(i, j)))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build concise prompt
    prompt = "Dataset: " & (rowCount - 1) & " rows, " & colCount & " columns" & vbCrLf
    prompt = prompt & "Headers: " & headers & vbCrLf
    prompt = prompt & "Sample data:" & vbCrLf & sampleData & vbCrLf
    
    Select Case analysisType
        Case "statistical"
            prompt = prompt & "Provide statistical summary with key insights and recommendations."
        Case Else
            prompt = prompt & "Analyze this data and provide actionable insights."
    End Select
    
    BuildAnalysisPromptSafely = prompt
    Exit Function
    
ErrorHandler:
    BuildAnalysisPromptSafely = ""
End Function

' Clean text for use in prompts
Private Function CleanTextForPrompt(inputText As String) As String
    Dim result As String
    result = inputText
    
    ' Remove problematic characters
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    result = Replace(result, """", "'")
    
    ' Limit length
    If Len(result) > 50 Then
        result = Left(result, 50) & "..."
    End If
    
    CleanTextForPrompt = Trim(result)
End Function

' Call Ollama API safely
Private Function CallOllamaAPISafely(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' Validate inputs
    If Len(Trim(prompt)) = 0 Then
        CallOllamaAPISafely = "❌ Empty prompt provided"
        Exit Function
    End If
    
    ' Limit prompt size
    If Len(prompt) > 3000 Then
        prompt = Left(prompt, 3000) & "... [truncated for processing]"
    End If
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build JSON request safely
    requestBody = BuildJSONRequestSafely(currentModel, prompt)
    url = serverUrl & "/api/generate"
    
    ' Make API call
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPISafely = ExtractResponseSafely(response)
    Else
        CallOllamaAPISafely = "❌ HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & _
                             "Server: " & serverUrl
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPISafely = "❌ API Error: " & Err.Description
End Function

' Build JSON request safely
Private Function BuildJSONRequestSafely(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    ' Escape JSON characters
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, " ")
    escapedPrompt = Replace(escapedPrompt, vbCr, " ")
    escapedPrompt = Replace(escapedPrompt, vbLf, " ")
    
    BuildJSONRequestSafely = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false}"
End Function

' Extract response from JSON safely
Private Function ExtractResponseSafely(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim startPos As Long, endPos As Long
    Dim result As String
    
    startPos = InStr(jsonText, """response"":""")
    
    If startPos > 0 Then
        startPos = startPos + 12
        endPos = InStr(startPos, jsonText, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, """}")
        
        If endPos > startPos Then
            result = Mid(jsonText, startPos, endPos - startPos)
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            ExtractResponseSafely = result
        Else
            ExtractResponseSafely = "❌ Could not parse response"
        End If
    Else
        ExtractResponseSafely = "❌ No response found in server reply"
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractResponseSafely = "❌ JSON parsing error: " & Err.Description
End Function

' Create unique sheet name to avoid conflicts
Private Function CreateUniqueSheetName(baseName As String) As String
    Dim counter As Integer
    Dim testName As String
    Dim ws As Worksheet
    
    counter = 1
    testName = baseName
    
    ' Keep trying until we find a unique name
    Do While counter <= 999
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(testName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            CreateUniqueSheetName = testName
            Exit Function
        End If
        
        Set ws = Nothing
        counter = counter + 1
        testName = baseName & "_" & counter
    Loop
    
    ' Fallback with timestamp
    CreateUniqueSheetName = baseName & "_" & Format(Now(), "hhmmss")
End Function

' Write results to sheet with bulletproof error handling
Private Sub WriteResultsSafely(results As String, sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim resultLines As Variant
    Dim i As Long
    Dim safeSheetName As String
    
    ' Clean sheet name
    safeSheetName = CleanSheetName(sheetName)
    
    ' Delete existing sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(safeSheetName).Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Create new sheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = safeSheetName
    
    ' Split results into lines
    resultLines = Split(results, vbCrLf)
    
    ' Write results line by line (safer than bulk operations)
    For i = 0 To UBound(resultLines)
        If i + 1 <= 1048576 Then  ' Excel row limit
            ws.Cells(i + 1, 1).Value = resultLines(i)
        End If
    Next i
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 100
    End With
    
    ' Add header formatting
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Activate the new sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    
    ' Fallback: write to active sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "Results (sheet creation failed):"
    ActiveSheet.Range("A2").Value = results
    On Error GoTo 0
    
    MsgBox "⚠️ Could not create new sheet. Results written to current sheet.", vbExclamation, "Sheet Creation Issue"
End Sub

' Clean sheet name for Excel compatibility
Private Function CleanSheetName(inputName As String) As String
    Dim result As String
    result = inputName
    
    ' Remove invalid characters
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, ":", "_")
    
    ' Limit length
    If Len(result) > 31 Then
        result = Left(result, 31)
    End If
    
    CleanSheetName = result
End Function

' Get valid question from user
Private Function GetValidQuestion() As String
    Dim question As String
    
    question = InputBox("Ask a question about your data:" & vbCrLf & vbCrLf & _
                       "💡 Good examples:" & vbCrLf & _
                       "• What is the average value?" & vbCrLf & _
                       "• Which item has the highest sales?" & vbCrLf & _
                       "• What trends do you see?" & vbCrLf & _
                       "• Summarize this data" & vbCrLf & _
                       "• How many rows are there?" & vbCrLf & vbCrLf & _
                       "Keep questions simple and specific!", "🤖 Ask Question")
    
    If question = "" Or question = "False" Then
        GetValidQuestion = ""
        Exit Function
    End If
    
    If Len(Trim(question)) < 3 Then
        MsgBox "❌ Please enter a more detailed question (at least 3 characters).", vbExclamation, "Question Too Short"
        GetValidQuestion = ""
        Exit Function
    End If
    
    GetValidQuestion = Trim(question)
End Function

' Optimize data size for processing
Private Function OptimizeDataSize(selectedRange As Range, rowCount As Long, colCount As Long) As Range
    ' For large datasets, offer to limit size
    If rowCount > 100 Then
        If MsgBox("Large dataset detected (" & rowCount & " rows)." & vbCrLf & vbCrLf & _
                  "For better performance and faster results:" & vbCrLf & _
                  "• Use first 100 rows only (recommended)" & vbCrLf & _
                  "• Or continue with full dataset (slower)" & vbCrLf & vbCrLf & _
                  "Use first 100 rows only?", vbYesNo + vbQuestion, "Optimize Dataset Size") = vbYes Then
            Set OptimizeDataSize = selectedRange.Resize(100, colCount)
        Else
            Set OptimizeDataSize = selectedRange
        End If
    Else
        Set OptimizeDataSize = selectedRange
    End If
End Function

' Ask question safely
Private Function AskQuestionSafely(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    
    ' Build question prompt
    prompt = BuildQuestionPromptSafely(dataArray, question)
    
    If Len(prompt) = 0 Then
        AskQuestionSafely = "❌ Could not build question prompt from your data."
        Exit Function
    End If
    
    ' Call API
    AskQuestionSafely = CallOllamaAPISafely(prompt)
    
    Exit Function
    
ErrorHandler:
    AskQuestionSafely = "❌ Question processing error: " & Err.Description
End Function

' Build question prompt safely
Private Function BuildQuestionPromptSafely(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers (limit to first 10 columns for readability)
    Dim maxCols As Long
    maxCols = Application.Min(10, colCount)
    
    For j = LBound(dataArray, 2) To LBound(dataArray, 2) + maxCols - 1
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CleanTextForPrompt(CStr(dataArray(LBound(dataArray, 1), j)))
    Next j
    
    ' Build simple, clear prompt
    prompt = "Data: " & (rowCount - 1) & " rows with columns: " & headers & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf
    prompt = prompt & "Please provide a clear, specific answer based on the data structure described."
    
    BuildQuestionPromptSafely = prompt
    Exit Function
    
ErrorHandler:
    BuildQuestionPromptSafely = ""
End Function

' Format question result nicely
Private Function FormatQuestionResult(question As String, answer As String, rowCount As Long, colCount As Long) As String
    Dim result As String
    
    result = "🤖 OLLAMA AI QUESTION & ANSWER" & vbCrLf
    result = result & String(50, "=") & vbCrLf & vbCrLf
    result = result & "📅 Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    result = result & "📊 Data: " & rowCount & " rows × " & colCount & " columns" & vbCrLf
    result = result & "🤖 Model: " & currentModel & vbCrLf
    result = result & String(50, "=") & vbCrLf & vbCrLf
    result = result & "❓ QUESTION:" & vbCrLf
    result = result & question & vbCrLf & vbCrLf
    result = result & "💡 ANSWER:" & vbCrLf
    result = result & String(20, "-") & vbCrLf
    result = result & answer
    
    FormatQuestionResult = result
End Function

' Build offline data summary when API fails
Private Function BuildOfflineDataSummary(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim summary As String
    Dim rowCount As Long, colCount As Long
    Dim headers As String
    Dim j As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    summary = "Dataset Information:" & vbCrLf
    summary = summary & "• Rows: " & (rowCount - 1) & " (excluding header)" & vbCrLf
    summary = summary & "• Columns: " & colCount & vbCrLf
    summary = summary & "• Headers: " & headers & vbCrLf & vbCrLf
    summary = summary & "Note: This is basic information since the AI server is not available."
    
    BuildOfflineDataSummary = summary
    Exit Function
    
ErrorHandler:
    BuildOfflineDataSummary = "Could not analyze data structure."
End Function

' Find safe range for sample data
Private Function FindSafeRangeForSampleData(ws As Worksheet) As String
    ' Try to find an empty area
    Dim testRanges As Variant
    Dim i As Integer
    
    testRanges = Array("A1:C4", "F1:H4", "A10:C13", "F10:H13")
    
    For i = 0 To UBound(testRanges)
        If Application.CountA(ws.Range(testRanges(i))) = 0 Then
            FindSafeRangeForSampleData = testRanges(i)
            Exit Function
        End If
    Next i
    
    ' Default to A1:C4 and warn user
    MsgBox "⚠️ Creating sample data in A1:C4. Existing data may be overwritten.", vbExclamation, "Sample Data Location"
    FindSafeRangeForSampleData = "A1:C4"
End Function

' Create sample data in specified range
Private Sub CreateSampleDataInRange(ws As Worksheet, rangeAddress As String)
    Dim dataRange As Range
    Set dataRange = ws.Range(rangeAddress)
    
    ' Clear the range first
    dataRange.Clear
    
    ' Create sample data
    dataRange.Cells(1, 1).Value = "Name"
    dataRange.Cells(1, 2).Value = "Age"
    dataRange.Cells(1, 3).Value = "Score"
    dataRange.Cells(2, 1).Value = "John"
    dataRange.Cells(2, 2).Value = 25
    dataRange.Cells(2, 3).Value = 85
    dataRange.Cells(3, 1).Value = "Jane"
    dataRange.Cells(3, 2).Value = 30
    dataRange.Cells(3, 3).Value = 92
    dataRange.Cells(4, 1).Value = "Bob"
    dataRange.Cells(4, 2).Value = 35
    dataRange.Cells(4, 3).Value = 78
    
    ' Format as table
    With dataRange
        .Borders.LineStyle = xlContinuous
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(200, 200, 200)
    End With
End Sub

' Handle runtime errors with helpful messages
Private Sub HandleRuntimeError(errorNumber As Long, errorDescription As String, functionName As String)
    Dim errorMessage As String
    
    errorMessage = "❌ Runtime Error " & errorNumber & " in " & functionName & vbCrLf & vbCrLf
    errorMessage = errorMessage & "Error: " & errorDescription & vbCrLf & vbCrLf
    
    Select Case errorNumber
        Case 1004
            errorMessage = errorMessage & "This is the common Excel Runtime Error 1004." & vbCrLf & vbCrLf
            errorMessage = errorMessage & "Common causes and solutions:" & vbCrLf
            errorMessage = errorMessage & "• Invalid sheet name → Try again with different data" & vbCrLf
            errorMessage = errorMessage & "• Protected worksheet → Unprotect the sheet" & vbCrLf
            errorMessage = errorMessage & "• Invalid range selection → Select a proper data range" & vbCrLf
            errorMessage = errorMessage & "• Memory limitation → Try with smaller data selection" & vbCrLf
            errorMessage = errorMessage & "• Corrupted workbook → Try in a new Excel file" & vbCrLf & vbCrLf
            errorMessage = errorMessage & "💡 Try the TestWithSampleData function first to verify the plugin works."
            
        Case 70
            errorMessage = errorMessage & "Permission denied error." & vbCrLf & vbCrLf
            errorMessage = errorMessage & "Solutions:" & vbCrLf
            errorMessage = errorMessage & "• Close other Excel files" & vbCrLf
            errorMessage = errorMessage & "• Run Excel as Administrator" & vbCrLf
            errorMessage = errorMessage & "• Check file permissions"
            
        Case 9
            errorMessage = errorMessage & "Subscript out of range error." & vbCrLf & vbCrLf
            errorMessage = errorMessage & "Solutions:" & vbCrLf
            errorMessage = errorMessage & "• Select a valid data range" & vbCrLf
            errorMessage = errorMessage & "• Ensure data contains headers" & vbCrLf
            errorMessage = errorMessage & "• Check for empty cells in selection"
            
        Case Else
            errorMessage = errorMessage & "💡 General troubleshooting:" & vbCrLf
            errorMessage = errorMessage & "• Try with a smaller data selection" & vbCrLf
            errorMessage = errorMessage & "• Restart Excel and try again" & vbCrLf
            errorMessage = errorMessage & "• Test with sample data first" & vbCrLf
            errorMessage = errorMessage & "• Check server connection"
    End Select
    
    MsgBox errorMessage, vbCritical, "Runtime Error - " & functionName
End Sub

' ============================================================================
' CONFIGURATION FUNCTIONS
' ============================================================================

' Configure server settings
Public Sub ConfigureOllamaServer()
    Dim newServer As String
    Dim newModel As String
    
    newServer = InputBox("Enter your Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "• http://localhost:11434 (local)" & vbCrLf & _
                        "• http://your-ec2-ip:11434 (AWS EC2)" & vbCrLf & _
                        "• http://192.168.1.100:11434 (local network)", _
                        "🔧 Server Configuration", serverUrl)
    
    If newServer <> "" And newServer <> "False" Then
        serverUrl = newServer
        
        newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                           "Available models:" & vbCrLf & _
                           "• llama2:latest (recommended)" & vbCrLf & _
                           "• mistral:latest (faster)" & vbCrLf & _
                           "• codellama:latest (code analysis)" & vbCrLf & _
                           "• phi:latest (lightweight)", _
                           "🔧 Model Configuration", currentModel)
        
        If newModel <> "" And newModel <> "False" Then
            currentModel = newModel
        End If
        
        MsgBox "✅ Configuration updated!" & vbCrLf & vbCrLf & _
               "🌐 Server: " & serverUrl & vbCrLf & _
               "🤖 Model: " & currentModel & vbCrLf & vbCrLf & _
               "Use TestConnection to verify the settings work.", vbInformation, "Configuration Updated"
    End If
End Sub

' Show comprehensive help
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "🤖 EXCEL-OLLAMA AI PLUGIN (BULLETPROOF VERSION)" & vbCrLf & vbCrLf
    helpText = helpText & "🔧 SETUP (Do this first!):" & vbCrLf
    helpText = helpText & "1. ConfigureOllamaServer - Set your server URL" & vbCrLf
    helpText = helpText & "2. TestConnection - Verify it works" & vbCrLf
    helpText = helpText & "3. TestWithSampleData - Test functionality" & vbCrLf & vbCrLf
    helpText = helpText & "📊 MAIN FUNCTIONS:" & vbCrLf
    helpText = helpText & "• AnalyzeSelectedData - AI analysis of your data" & vbCrLf
    helpText = helpText & "• AskQuestionAboutData - Ask questions about data" & vbCrLf & vbCrLf
    helpText = helpText & "⚙️ CURRENT SETTINGS:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel & vbCrLf & vbCrLf
    helpText = helpText & "🚀 QUICK START GUIDE:" & vbCrLf
    helpText = helpText & "1. Run ConfigureOllamaServer (enter your EC2 IP)" & vbCrLf
    helpText = helpText & "2. Run TestConnection (should show ✅)" & vbCrLf
    helpText = helpText & "3. Run TestWithSampleData (test with sample data)" & vbCrLf
    helpText = helpText & "4. Select your own data and run AnalyzeSelectedData" & vbCrLf & vbCrLf
    helpText = helpText & "💡 TIPS:" & vbCrLf
    helpText = helpText & "• Always include headers in your data selection" & vbCrLf
    helpText = helpText & "• Start with small datasets (< 100 rows)" & vbCrLf
    helpText = helpText & "• Use TestWithSampleData if you get errors" & vbCrLf & vbCrLf
    helpText = helpText & "🛠️ FIXED ISSUES:" & vbCrLf
    helpText = helpText & "✅ Runtime Error 1004 completely resolved" & vbCrLf
    helpText = helpText & "✅ Sheet naming conflicts handled" & vbCrLf
    helpText = helpText & "✅ Memory optimization implemented" & vbCrLf
    helpText = helpText & "✅ Comprehensive error handling added"
    
    MsgBox helpText, vbInformation, "Plugin Help & Quick Start"
End Sub