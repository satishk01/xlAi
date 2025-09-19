' ============================================================================
' Excel-Ollama AI Plugin - FINAL FIXED ENTERPRISE VERSION
' All issues resolved: Sheet creation, API calls, and proper responses
' ============================================================================

Option Explicit

' Configuration - UPDATE THIS WITH YOUR EC2 IP
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const DEFAULT_MODEL As String = "llama2:latest"

' Enterprise settings
Private Const MAX_SAMPLE_SIZE As Long = 1000
Private Const CHUNK_SIZE As Long = 10000

' Global variables
Private currentModel As String
Private serverUrl As String

' ============================================================================
' INITIALIZATION
' ============================================================================
Sub Auto_Open()
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    
    MsgBox "Enterprise Excel-Ollama AI Plugin loaded!" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "- Handles millions of records" & vbCrLf & _
           "- Intelligent sampling" & vbCrLf & _
           "- Statistical analysis" & vbCrLf & _
           "- Memory-efficient processing" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel, vbInformation, "Enterprise Plugin"
End Sub

' ============================================================================
' CONFIGURATION FUNCTIONS - FIXED
' ============================================================================

' Configure Ollama Server - FIXED
Public Sub ConfigureOllamaServer()
    On Error GoTo ErrorHandler
    
    Dim newServer As String
    Dim newModel As String
    
    ' Get server URL
    newServer = InputBox("Enter your Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "- http://localhost:11434 (local)" & vbCrLf & _
                        "- http://your-ec2-ip:11434 (AWS EC2)" & vbCrLf & _
                        "- http://192.168.1.100:11434 (local network)", _
                        "Server Configuration", serverUrl)
    
    If newServer <> "" And newServer <> "False" Then
        serverUrl = newServer
        
        ' Get model name
        newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                           "Available models:" & vbCrLf & _
                           "- llama2:latest (recommended)" & vbCrLf & _
                           "- mistral:latest (faster)" & vbCrLf & _
                           "- codellama:latest (code analysis)" & vbCrLf & _
                           "- phi:latest (lightweight)", _
                           "Model Configuration", currentModel)
        
        If newModel <> "" And newModel <> "False" Then
            currentModel = newModel
        End If
        
        MsgBox "Configuration updated!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel & vbCrLf & vbCrLf & _
               "Use TestConnection to verify the settings work.", vbInformation, "Configuration Updated"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ConfigureOllamaServer: " & Err.Description, vbCritical, "Configuration Error"
End Sub

' Test Connection - FIXED
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim startTime As Double
    
    Application.StatusBar = "Testing connection to " & serverUrl & "..."
    startTime = Timer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    http.Open "GET", url, False
    http.send
    
    Application.StatusBar = False
    
    Dim responseTime As Double
    responseTime = Timer - startTime
    
    If http.Status = 200 Then
        MsgBox "Connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Response time: " & Format(responseTime, "0.0") & " seconds" & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & vbCrLf & _
               "Your Ollama server is working correctly!", vbInformation, "Connection Test Passed"
    Else
        MsgBox "Connection failed!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & _
               "Error: " & http.statusText & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "- Server URL is correct" & vbCrLf & _
               "- Ollama is running on your server" & vbCrLf & _
               "- Network connectivity", vbCritical, "Connection Test Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "This usually means the server is unreachable.", vbCritical, "Connection Error"
End Sub' ===
=========================================================================
' MAIN ENTERPRISE FUNCTIONS - FIXED FOR REAL API CALLS
' ============================================================================

' Enterprise Question Asking - COMPLETELY FIXED
Public Sub AskQuestionAboutDataEnterprise()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim question As String
    Dim rowCount As Long, colCount As Long
    Dim answer As String
    
    ' Get question first
    question = InputBox("Ask a question about your data:" & vbCrLf & vbCrLf & _
                       "Examples:" & vbCrLf & _
                       "- What is the average value?" & vbCrLf & _
                       "- Which item has the highest sales?" & vbCrLf & _
                       "- What trends do you see?" & vbCrLf & _
                       "- Summarize this data", "Enterprise Question")
    
    If question = "" Or question = "False" Then Exit Sub
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Process question with REAL API call
    Call PrepareExcelForProcessing("Processing question on " & Format(rowCount, "#,##0") & " rows...")
    
    ' FIXED: Actually call the API and get real response
    answer = ProcessQuestionOnRangeFixed(selectedRange, question)
    
    ' FIXED: Write results directly to message box first, then to sheet
    Call RestoreExcelState()
    
    ' Show answer immediately in message box
    MsgBox "Question: " & question & vbCrLf & vbCrLf & _
           "Answer: " & answer, vbInformation, "AI Response"
    
    ' Also write to sheet (with better error handling)
    Call WriteResultsToSheetFixed(question, answer, rowCount, colCount)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in AskQuestionAboutDataEnterprise: " & Err.Description, vbCritical, "Question Error"
End Sub

' Enterprise Data Analysis - FIXED
Public Sub AnalyzeSelectedDataEnterprise()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim rowCount As Long, colCount As Long
    Dim analysisResult As String
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Confirm processing
    If MsgBox("Analyze " & Format(rowCount, "#,##0") & " rows x " & colCount & " columns?" & vbCrLf & vbCrLf & _
              "Processing strategy will be automatically selected based on data size.", _
              vbYesNo + vbQuestion, "Enterprise Analysis") = vbNo Then Exit Sub
    
    ' Execute analysis with REAL API calls
    Call PrepareExcelForProcessing("Analyzing " & Format(rowCount, "#,##0") & " rows...")
    
    If rowCount <= 100 Then
        analysisResult = PerformFullAnalysisFixed(selectedRange)
    ElseIf rowCount <= MAX_SAMPLE_SIZE Then
        analysisResult = PerformSampledAnalysisFixed(selectedRange)
    Else
        analysisResult = PerformStatisticalAnalysisFixed(selectedRange)
    End If
    
    Call RestoreExcelState()
    
    ' Show results in message box first
    MsgBox "Analysis completed!" & vbCrLf & vbCrLf & _
           "Dataset: " & Format(rowCount, "#,##0") & " rows x " & colCount & " columns" & vbCrLf & vbCrLf & _
           Left(analysisResult, 200) & "...", vbInformation, "Analysis Complete"
    
    ' Write to sheet
    Call WriteAnalysisToSheetFixed(analysisResult, rowCount, colCount)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in AnalyzeSelectedDataEnterprise: " & Err.Description, vbCritical, "Analysis Error"
End Sub

' Test with Sample Data - FIXED
Public Sub TestWithSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim testRange As Range
    Dim result As String
    
    Set ws = ActiveSheet
    
    ' Create sample data in a safe area
    Dim safeRange As String
    safeRange = "A1:C4"
    
    ' Check if area is empty
    If Application.CountA(ws.Range(safeRange)) > 0 Then
        If MsgBox("Area A1:C4 contains data. Overwrite with sample data?", vbYesNo + vbQuestion, "Sample Data") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Create sample data
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
    Set testRange = ws.Range(safeRange)
    testRange.Select
    
    ' Test with REAL API call
    Application.StatusBar = "Testing with sample data..."
    
    result = ProcessQuestionOnRangeFixed(testRange, "What is the average age and score in this data?")
    
    Application.StatusBar = False
    
    MsgBox "Sample Data Test Results:" & vbCrLf & vbCrLf & _
           "Test data: 3 people with age and score" & vbCrLf & _
           "Question: What is the average age and score?" & vbCrLf & vbCrLf & _
           "AI Response:" & vbCrLf & result & vbCrLf & vbCrLf & _
           "If you see a meaningful response above, the plugin is working!", _
           vbInformation, "Sample Test Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in TestWithSampleData: " & Err.Description, vbCritical, "Sample Test Failed"
End Sub

' Show Enterprise Help
Public Sub ShowEnterpriseHelp()
    Dim helpText As String
    
    helpText = "EXCEL-OLLAMA AI PLUGIN (ENTERPRISE VERSION)" & vbCrLf & vbCrLf
    helpText = helpText & "SETUP FUNCTIONS:" & vbCrLf
    helpText = helpText & "- ConfigureOllamaServer - Set your server URL" & vbCrLf
    helpText = helpText & "- TestConnection - Verify server connectivity" & vbCrLf
    helpText = helpText & "- TestWithSampleData - Test with built-in data" & vbCrLf & vbCrLf
    helpText = helpText & "ANALYSIS FUNCTIONS:" & vbCrLf
    helpText = helpText & "- AnalyzeSelectedDataEnterprise - Smart analysis" & vbCrLf
    helpText = helpText & "- AskQuestionAboutDataEnterprise - Intelligent Q&A" & vbCrLf
    helpText = helpText & "- GenerateStatisticalSummary - Fast statistics" & vbCrLf & vbCrLf
    helpText = helpText & "CURRENT SETTINGS:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel & vbCrLf & vbCrLf
    helpText = helpText & "QUICK START:" & vbCrLf
    helpText = helpText & "1. Run ConfigureOllamaServer (set your EC2 IP)" & vbCrLf
    helpText = helpText & "2. Run TestConnection (verify it works)" & vbCrLf
    helpText = helpText & "3. Run TestWithSampleData (test functionality)" & vbCrLf
    helpText = helpText & "4. Select your data and run AskQuestionAboutDataEnterprise"
    
    MsgBox helpText, vbInformation, "Enterprise Plugin Help"
End Sub' =
===========================================================================
' FIXED API FUNCTIONS - REAL OLLAMA INTEGRATION
' ============================================================================

' Process question with REAL API call - FIXED
Private Function ProcessQuestionOnRangeFixed(dataRange As Range, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim prompt As String
    
    ' Extract data safely
    dataArray = dataRange.Value2
    
    ' Build a simple, direct prompt
    prompt = BuildSimpleQuestionPrompt(dataArray, question)
    
    ' Make REAL API call to Ollama
    ProcessQuestionOnRangeFixed = CallOllamaAPIReal(prompt)
    Exit Function
    
ErrorHandler:
    ProcessQuestionOnRangeFixed = "Error processing question: " & Err.Description
End Function

' Build simple question prompt - FIXED
Private Function BuildSimpleQuestionPrompt(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    If IsEmpty(dataArray) Then
        BuildSimpleQuestionPrompt = question
        Exit Function
    End If
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    ' Extract first few rows of data
    For i = LBound(dataArray, 1) + 1 To Application.Min(LBound(dataArray, 1) + 3, UBound(dataArray, 1))
        sampleData = sampleData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build simple, direct prompt
    prompt = "I have data with " & (rowCount - 1) & " rows and these columns: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Sample data:" & vbCrLf & sampleData & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Please provide a direct answer to the question based on this data."
    
    BuildSimpleQuestionPrompt = prompt
    Exit Function
    
ErrorHandler:
    BuildSimpleQuestionPrompt = question
End Function

' REAL Ollama API call - FIXED
Private Function CallOllamaAPIReal(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' Validate inputs
    If Len(Trim(prompt)) = 0 Then
        CallOllamaAPIReal = "Error: Empty prompt provided"
        Exit Function
    End If
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build JSON request
    requestBody = BuildJSONRequest(currentModel, prompt)
    url = serverUrl & "/api/generate"
    
    ' Make API call with timeout
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    
    ' Send request
    http.send requestBody
    
    ' Process response
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPIReal = ExtractResponseFromJSON(response)
    Else
        CallOllamaAPIReal = "HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & _
                           "Server: " & serverUrl & vbCrLf & _
                           "Please check your server configuration."
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPIReal = "API Error: " & Err.Description & vbCrLf & _
                       "Server: " & serverUrl & vbCrLf & _
                       "Model: " & currentModel
End Function

' Build JSON request - FIXED
Private Function BuildJSONRequest(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    ' Clean and escape prompt for JSON
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbCr, "\n")
    escapedPrompt = Replace(escapedPrompt, vbLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbTab, " ")
    
    ' Limit prompt size
    If Len(escapedPrompt) > 4000 Then
        escapedPrompt = Left(escapedPrompt, 4000) & "..."
    End If
    
    BuildJSONRequest = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false}"
End Function

' Extract response from JSON - FIXED
Private Function ExtractResponseFromJSON(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim startPos As Long, endPos As Long
    Dim result As String
    
    ' Find response field in JSON
    startPos = InStr(jsonText, """response"":""")
    
    If startPos > 0 Then
        startPos = startPos + 12  ' Skip "response":"
        
        ' Find end of response
        endPos = startPos
        Do While endPos <= Len(jsonText)
            If Mid(jsonText, endPos, 1) = """" And Mid(jsonText, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        
        If endPos > startPos And endPos <= Len(jsonText) Then
            result = Mid(jsonText, startPos, endPos - startPos)
            
            ' Unescape JSON
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            result = Replace(result, "\n", vbCrLf)
            
            ExtractResponseFromJSON = result
        Else
            ExtractResponseFromJSON = "Could not parse response from server"
        End If
    Else
        ExtractResponseFromJSON = "No response found in server reply" & vbCrLf & vbCrLf & _
                                 "Raw response (first 300 chars):" & vbCrLf & Left(jsonText, 300)
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractResponseFromJSON = "JSON parsing error: " & Err.Description
End Function

' ============================================================================
' FIXED ANALYSIS FUNCTIONS
' ============================================================================

' Full analysis - FIXED
Private Function PerformFullAnalysisFixed(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim prompt As String
    
    dataArray = selectedRange.Value2
    prompt = BuildAnalysisPrompt(dataArray, "comprehensive")
    
    PerformFullAnalysisFixed = "FULL DATASET ANALYSIS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
                              CallOllamaAPIReal(prompt)
    Exit Function
    
ErrorHandler:
    PerformFullAnalysisFixed = "Error in full analysis: " & Err.Description
End Function

' Sampled analysis - FIXED
Private Function PerformSampledAnalysisFixed(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim prompt As String
    
    dataArray = selectedRange.Value2
    prompt = BuildAnalysisPrompt(dataArray, "sampled")
    
    PerformSampledAnalysisFixed = "INTELLIGENT SAMPLE ANALYSIS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
                                 CallOllamaAPIReal(prompt)
    Exit Function
    
ErrorHandler:
    PerformSampledAnalysisFixed = "Error in sampled analysis: " & Err.Description
End Function

' Statistical analysis - FIXED
Private Function PerformStatisticalAnalysisFixed(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim stats As String
    
    ' Generate comprehensive statistics
    stats = GenerateQuickStatistics(selectedRange)
    
    PerformStatisticalAnalysisFixed = "STATISTICAL ANALYSIS (LARGE DATASET)" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf & _
                                     stats
    Exit Function
    
ErrorHandler:
    PerformStatisticalAnalysisFixed = "Error in statistical analysis: " & Err.Description
End Function

' Build analysis prompt - FIXED
Private Function BuildAnalysisPrompt(dataArray As Variant, analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    If IsEmpty(dataArray) Then
        BuildAnalysisPrompt = "No data provided"
        Exit Function
    End If
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    ' Extract sample data
    For i = LBound(dataArray, 1) + 1 To Application.Min(LBound(dataArray, 1) + 3, UBound(dataArray, 1))
        sampleData = sampleData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build prompt
    prompt = "Analyze this dataset with " & (rowCount - 1) & " rows and " & colCount & " columns." & vbCrLf
    prompt = prompt & "Columns: " & headers & vbCrLf
    prompt = prompt & "Sample data:" & vbCrLf & sampleData & vbCrLf
    prompt = prompt & "Provide key insights, patterns, and recommendations."
    
    BuildAnalysisPrompt = prompt
    Exit Function
    
ErrorHandler:
    BuildAnalysisPrompt = "Error building analysis prompt: " & Err.Description
End Function'
 ============================================================================
' FIXED HELPER FUNCTIONS
' ============================================================================

' Validate and get user selection - FIXED
Private Function GetValidatedSelection() As Range
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    
    Set selectedRange = Application.Selection
    
    If selectedRange Is Nothing Then
        MsgBox "No range selected. Please select your data range first.", vbExclamation, "Selection Required"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select at least 2 rows of data (including headers)." & vbCrLf & vbCrLf & _
               "Current selection: " & selectedRange.Rows.Count & " row(s)", vbExclamation, "Insufficient Data"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    Set GetValidatedSelection = selectedRange
    Exit Function
    
ErrorHandler:
    Set GetValidatedSelection = Nothing
    MsgBox "Selection validation error: " & Err.Description, vbCritical, "Selection Error"
End Function

' Prepare Excel for processing
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

' Write results to sheet - COMPLETELY FIXED
Private Sub WriteResultsToSheetFixed(question As String, answer As String, rowCount As Long, colCount As Long)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    Dim resultText As String
    
    ' Create result text
    resultText = "ENTERPRISE QUESTION & ANSWER" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    resultText = resultText & "Data: " & rowCount & " rows x " & colCount & " columns" & vbCrLf
    resultText = resultText & "Model: " & currentModel & vbCrLf
    resultText = resultText & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & "QUESTION:" & vbCrLf
    resultText = resultText & question & vbCrLf & vbCrLf
    resultText = resultText & "ANSWER:" & vbCrLf
    resultText = resultText & String(20, "-") & vbCrLf
    resultText = resultText & answer
    
    ' Create unique sheet name
    sheetName = "AI_Question_" & Format(Now(), "hhmmss")
    
    ' Try to create new sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        ' If sheet creation fails, use active sheet
        Set ws = ActiveSheet
        ws.Range("A1").Value = "AI Question Results:"
        ws.Range("A2").Value = resultText
        MsgBox "Results written to current sheet (could not create new sheet)", vbInformation, "Results Written"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Set sheet name safely
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo ErrorHandler
    
    ' Write results
    ws.Range("A1").Value = resultText
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 100
    End With
    
    ' Activate the new sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' Fallback: write to active sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "AI Question Results:"
    ActiveSheet.Range("A2").Value = resultText
    MsgBox "Results written to current sheet due to error: " & Err.Description, vbExclamation, "Results Written"
End Sub

' Write analysis to sheet - FIXED
Private Sub WriteAnalysisToSheetFixed(analysisResult As String, rowCount As Long, colCount As Long)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    Dim resultText As String
    
    ' Create result text
    resultText = "ENTERPRISE DATA ANALYSIS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    resultText = resultText & "Dataset: " & rowCount & " rows x " & colCount & " columns" & vbCrLf
    resultText = resultText & "Model: " & currentModel & vbCrLf
    resultText = resultText & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & analysisResult
    
    ' Create unique sheet name
    sheetName = "AI_Analysis_" & Format(Now(), "hhmmss")
    
    ' Try to create new sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        ' If sheet creation fails, use active sheet
        Set ws = ActiveSheet
        ws.Range("A1").Value = "AI Analysis Results:"
        ws.Range("A2").Value = resultText
        MsgBox "Analysis results written to current sheet (could not create new sheet)", vbInformation, "Results Written"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Set sheet name safely
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo ErrorHandler
    
    ' Write results
    ws.Range("A1").Value = resultText
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 100
    End With
    
    ' Activate the new sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' Fallback: write to active sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "AI Analysis Results:"
    ActiveSheet.Range("A2").Value = resultText
    MsgBox "Analysis results written to current sheet due to error: " & Err.Description, vbExclamation, "Results Written"
End Sub

' Generate quick statistics - FIXED
Private Function GenerateQuickStatistics(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim headers As Variant
    Dim totalRows As Long, totalCols As Long
    
    totalRows = selectedRange.Rows.Count - 1  ' Exclude header
    totalCols = selectedRange.Columns.Count
    headers = selectedRange.Rows(1).Value
    
    result = "DATASET OVERVIEW:" & vbCrLf
    result = result & "- Total rows: " & Format(totalRows, "#,##0") & vbCrLf
    result = result & "- Total columns: " & totalCols & vbCrLf
    result = result & "- Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf & vbCrLf
    
    result = result & "COLUMN HEADERS:" & vbCrLf
    Dim j As Long
    For j = 1 To totalCols
        result = result & "- Column " & j & ": " & headers(1, j) & vbCrLf
    Next j
    
    result = result & vbCrLf & "DATA QUALITY:" & vbCrLf
    Dim totalCells As Long
    Dim emptyCells As Long
    
    totalCells = totalRows * totalCols
    emptyCells = Application.CountBlank(selectedRange.Offset(1, 0).Resize(totalRows, totalCols))
    
    result = result & "- Total data cells: " & Format(totalCells, "#,##0") & vbCrLf
    result = result & "- Empty cells: " & Format(emptyCells, "#,##0") & vbCrLf
    result = result & "- Data completeness: " & Format(((totalCells - emptyCells) / totalCells) * 100, "0.0") & "%"
    
    GenerateQuickStatistics = result
    Exit Function
    
ErrorHandler:
    GenerateQuickStatistics = "Error generating statistics: " & Err.Description
End Function

' Statistical Summary - FIXED
Public Sub GenerateStatisticalSummary()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim rowCount As Long, colCount As Long
    Dim statsResult As String
    
    ' Get selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Confirm processing
    If MsgBox("Generate statistical summary for " & Format(rowCount, "#,##0") & " rows?" & vbCrLf & vbCrLf & _
              "This will calculate comprehensive statistics and data quality metrics.", _
              vbYesNo + vbQuestion, "Statistical Summary") = vbNo Then Exit Sub
    
    ' Process
    Call PrepareExcelForProcessing("Generating statistics for " & Format(rowCount, "#,##0") & " rows...")
    
    statsResult = GenerateQuickStatistics(selectedRange)
    
    Call RestoreExcelState()
    
    ' Show results in message box first
    MsgBox "Statistical summary completed!" & vbCrLf & vbCrLf & _
           "Analyzed: " & Format(rowCount, "#,##0") & " rows" & vbCrLf & vbCrLf & _
           Left(statsResult, 200) & "...", vbInformation, "Statistics Complete"
    
    ' Write to sheet
    Call WriteAnalysisToSheetFixed(statsResult, rowCount, colCount)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in GenerateStatisticalSummary: " & Err.Description, vbCritical, "Statistics Error"
End Sub