' Instructions to create Excel-Ollama AI Plugin (.xlam file)
' 
' STEP 1: Create a new Excel workbook
' STEP 2: Press Alt+F11 to open VBA Editor
' STEP 3: Insert > Module and paste the code below
' STEP 4: Save as Excel Add-in (.xlam) format
' STEP 5: Install the add-in in Excel

' ============================================================================
' Excel-Ollama AI Plugin - VBA Add-in
' No Python required on Windows!
' ============================================================================

Option Explicit

' Configuration - UPDATE THIS WITH YOUR EC2 IP
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const WEB_SERVER As String = "http://YOUR_EC2_IP:3000"
Private Const DEFAULT_MODEL As String = "llama2:latest"

' Global variables
Private currentModel As String
Private serverUrl As String

' Initialize add-in
Private Sub Workbook_Open()
    ' Set default values
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    
    ' Add custom ribbon (if using Excel 2007+)
    Call AddCustomRibbon
    
    ' Show welcome message
    MsgBox "🤖 Excel-Ollama AI Plugin loaded successfully!" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel & vbCrLf & vbCrLf & _
           "Use Alt+F8 to run macros or check the Add-ins tab.", vbInformation, "Ollama AI Plugin"
End Sub

' ============================================================================
' MAIN ANALYSIS FUNCTIONS
' ============================================================================

' Analyze selected data
Public Sub AnalyzeSelectedData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range with at least 2 rows (including headers)", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Analyzing data with Ollama AI..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Perform analysis
    analysisResult = CallOllamaAPI(dataArray, "statistical")
    
    ' Write results to new sheet
    Call WriteResultsToSheet(analysisResult, "AI_Analysis_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Analysis completed! Check the 'AI_Analysis_Results' sheet.", vbInformation, "Ollama AI"
End Sub

' Ask question about data
Public Sub AskQuestionAboutData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim question As String
    Dim answer As String
    
    ' Get user question
    question = InputBox("Ask a question about your data:" & vbCrLf & vbCrLf & _
                       "Examples:" & vbCrLf & _
                       "• What trends do you see?" & vbCrLf & _
                       "• Are there any outliers?" & vbCrLf & _
                       "• What patterns exist?", "🤖 Ollama AI Query")
    
    If question = "" Then Exit Sub
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range first", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Processing your question with Ollama AI..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Ask question
    answer = AskOllamaQuestion(dataArray, question)
    
    ' Write results to new sheet
    Call WriteResultsToSheet("QUESTION: " & question & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
                            "ANSWER: " & vbCrLf & answer, "AI_Query_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Query completed! Check the 'AI_Query_Results' sheet.", vbInformation, "Ollama AI"
End Sub

' Generate comprehensive report
Public Sub GenerateComprehensiveReport()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim report As String
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range for report generation", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Generating comprehensive AI report..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Generate report with multiple analyses
    report = GenerateMultiAnalysisReport(dataArray)
    
    ' Write results to new sheet
    Call WriteResultsToSheet(report, "AI_Comprehensive_Report")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Comprehensive report generated! Check the 'AI_Comprehensive_Report' sheet.", vbInformation, "Ollama AI"
End Sub

' Trend analysis
Public Sub AnalyzeTrends()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim trendResult As String
    
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range with at least 2 rows (including headers)", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    Application.StatusBar = "🤖 Analyzing trends with Ollama AI..."
    Application.ScreenUpdating = False
    
    dataArray = selectedRange.Value
    trendResult = CallOllamaAPI(dataArray, "trends")
    
    Call WriteResultsToSheet(trendResult, "AI_Trend_Analysis")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Trend analysis completed! Check the 'AI_Trend_Analysis' sheet.", vbInformation, "Ollama AI"
End Sub

' Pattern detection
Public Sub DetectPatterns()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim patternResult As String
    
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range with at least 2 rows (including headers)", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    Application.StatusBar = "🤖 Detecting patterns with Ollama AI..."
    Application.ScreenUpdating = False
    
    dataArray = selectedRange.Value
    patternResult = CallOllamaAPI(dataArray, "patterns")
    
    Call WriteResultsToSheet(patternResult, "AI_Pattern_Detection")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Pattern detection completed! Check the 'AI_Pattern_Detection' sheet.", vbInformation, "Ollama AI"
End Sub

' ============================================================================
' CONFIGURATION FUNCTIONS
' ============================================================================

' Configure Ollama server
Public Sub ConfigureOllamaServer()
    Dim newServer As String
    Dim newModel As String
    
    ' Get server URL
    newServer = InputBox("Enter Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "• http://localhost:11434 (local)" & vbCrLf & _
                        "• http://your-ec2-ip:11434 (EC2)", _
                        "Server Configuration", serverUrl)
    
    If newServer = "" Then Exit Sub
    
    ' Get model name
    newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                       "Available models:" & vbCrLf & _
                       "• llama2:latest (recommended)" & vbCrLf & _
                       "• mistral:latest (faster)" & vbCrLf & _
                       "• codellama:latest (code analysis)" & vbCrLf & _
                       "• phi:latest (lightweight)", _
                       "Model Configuration", currentModel)
    
    If newModel = "" Then Exit Sub
    
    ' Update configuration
    serverUrl = newServer
    currentModel = newModel
    
    ' Test connection
    If TestConnection() Then
        MsgBox "✅ Configuration updated and connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel, vbInformation, "Ollama AI"
    Else
        MsgBox "⚠️ Configuration updated but connection failed." & vbCrLf & _
               "Please check your server URL and ensure Ollama is running.", vbExclamation, "Ollama AI"
    End If
End Sub

' Test connection to Ollama server
Public Sub TestOllamaConnection()
    If TestConnection() Then
        MsgBox "✅ Connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel, vbInformation, "Ollama AI"
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "• Server URL: " & serverUrl & vbCrLf & _
               "• Ollama is running" & vbCrLf & _
               "• Network connectivity", vbCritical, "Ollama AI"
    End If
End Sub

' ============================================================================
' CORE API FUNCTIONS
' ============================================================================

' Main function to call Ollama API
Private Function CallOllamaAPI(dataArray As Variant, analysisType As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim prompt As String
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build prompt based on analysis type
    prompt = BuildAnalysisPrompt(dataArray, analysisType)
    
    ' Build request body (escape JSON properly)
    requestBody = "{""model"":""" & currentModel & """,""prompt"":""" & EscapeJSON(prompt) & """,""stream"":false}"
    
    ' Make API call
    url = serverUrl & "/api/generate"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPI = ExtractResponseFromJSON(response)
    Else
        CallOllamaAPI = "❌ Error: HTTP " & http.Status & " - " & http.statusText & vbCrLf & vbCrLf & _
                       "Please check your server configuration and try again."
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPI = "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
                   "Server: " & serverUrl & vbCrLf & _
                   "Please ensure Ollama is running and accessible."
End Function

' Ask question function
Private Function AskOllamaQuestion(dataArray As Variant, question As String) As String
    Dim prompt As String
    
    ' Build question prompt
    prompt = BuildQuestionPrompt(dataArray, question)
    
    ' Use the main API function with custom prompt
    AskOllamaQuestion = CallOllamaAPIWithPrompt(prompt)
End Function

' Call API with custom prompt
Private Function CallOllamaAPIWithPrompt(prompt As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    requestBody = "{""model"":""" & currentModel & """,""prompt"":""" & EscapeJSON(prompt) & """,""stream"":false}"
    url = serverUrl & "/api/generate"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPIWithPrompt = ExtractResponseFromJSON(response)
    Else
        CallOllamaAPIWithPrompt = "❌ Error: HTTP " & http.Status & " - " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPIWithPrompt = "❌ Connection Error: " & Err.Description
End Function

' Generate comprehensive report with multiple analyses
Private Function GenerateMultiAnalysisReport(dataArray As Variant) As String
    Dim report As String
    Dim statisticalAnalysis As String
    Dim trendAnalysis As String
    Dim patternAnalysis As String
    
    ' Header
    report = "🤖 COMPREHENSIVE DATA ANALYSIS REPORT" & vbCrLf
    report = report & String(60, "=") & vbCrLf
    report = report & "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    report = report & "Model: " & currentModel & vbCrLf
    report = report & "Data Size: " & (UBound(dataArray, 1) - 1) & " rows × " & UBound(dataArray, 2) & " columns" & vbCrLf
    report = report & String(60, "=") & vbCrLf & vbCrLf
    
    ' Statistical Analysis
    Application.StatusBar = "🤖 Performing statistical analysis..."
    statisticalAnalysis = CallOllamaAPI(dataArray, "statistical")
    report = report & "📊 STATISTICAL ANALYSIS" & vbCrLf
    report = report & String(30, "-") & vbCrLf
    report = report & statisticalAnalysis & vbCrLf & vbCrLf
    
    ' Trend Analysis
    Application.StatusBar = "🤖 Analyzing trends..."
    trendAnalysis = CallOllamaAPI(dataArray, "trends")
    report = report & "📈 TREND ANALYSIS" & vbCrLf
    report = report & String(30, "-") & vbCrLf
    report = report & trendAnalysis & vbCrLf & vbCrLf
    
    ' Pattern Analysis
    Application.StatusBar = "🤖 Detecting patterns..."
    patternAnalysis = CallOllamaAPI(dataArray, "patterns")
    report = report & "🔍 PATTERN DETECTION" & vbCrLf
    report = report & String(30, "-") & vbCrLf
    report = report & patternAnalysis & vbCrLf & vbCrLf
    
    ' Summary
    report = report & "📋 EXECUTIVE SUMMARY" & vbCrLf
    report = report & String(30, "-") & vbCrLf
    report = report & "This comprehensive analysis examined your data from multiple perspectives:" & vbCrLf
    report = report & "• Statistical characteristics and data quality" & vbCrLf
    report = report & "• Trends and temporal patterns" & vbCrLf
    report = report & "• Anomalies and hidden patterns" & vbCrLf & vbCrLf
    report = report & "Review each section above for detailed insights and recommendations."
    
    GenerateMultiAnalysisReport = report
End Function

' Test connection function
Private Function TestConnection() As Boolean
    Dim http As Object
    Dim url As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    On Error GoTo ErrorHandler
    
    http.Open "GET", url, False
    http.send
    
    TestConnection = (http.Status = 200)
    Exit Function
    
ErrorHandler:
    TestConnection = False
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' Build analysis prompt based on data and type
Private Function BuildAnalysisPrompt(dataArray As Variant, analysisType As String) As String
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Integer, j As Integer
    Dim rowCount As Integer, colCount As Integer
    
    rowCount = UBound(dataArray, 1) - 1  ' Subtract 1 for header row
    colCount = UBound(dataArray, 2)
    
    ' Extract headers
    For j = 1 To colCount
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(dataArray(1, j))
    Next j
    
    ' Extract sample data (first 5 rows)
    For i = 2 To Application.Min(6, UBound(dataArray, 1))
        For j = 1 To colCount
            If j > 1 Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build base prompt
    prompt = "Analyze this dataset with " & rowCount & " rows and " & colCount & " columns." & vbCrLf & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Sample data (first 5 rows):" & vbCrLf & sampleData & vbCrLf
    
    ' Add analysis-specific instructions
    Select Case analysisType
        Case "statistical"
            prompt = prompt & "Provide a comprehensive statistical analysis including:" & vbCrLf
            prompt = prompt & "1. Summary statistics for numeric columns" & vbCrLf
            prompt = prompt & "2. Data quality assessment (missing values, outliers)" & vbCrLf
            prompt = prompt & "3. Key patterns and correlations between variables" & vbCrLf
            prompt = prompt & "4. Business insights and actionable recommendations" & vbCrLf
            prompt = prompt & "Format your response with clear sections and bullet points."
            
        Case "trends"
            prompt = prompt & "Analyze trends and patterns in this data:" & vbCrLf
            prompt = prompt & "1. Identify time-based trends if applicable" & vbCrLf
            prompt = prompt & "2. Detect increasing, decreasing, or cyclical patterns" & vbCrLf
            prompt = prompt & "3. Seasonal or periodic variations" & vbCrLf
            prompt = prompt & "4. Forecast insights and future predictions" & vbCrLf
            prompt = prompt & "Provide actionable business recommendations based on trends."
            
        Case "patterns"
            prompt = prompt & "Detect patterns and anomalies in this data:" & vbCrLf
            prompt = prompt & "1. Identify unusual values or outliers" & vbCrLf
            prompt = prompt & "2. Find correlations and relationships between variables" & vbCrLf
            prompt = prompt & "3. Discover clustering or grouping patterns" & vbCrLf
            prompt = prompt & "4. Assess data quality issues" & vbCrLf
            prompt = prompt & "Highlight the most important findings and their business implications."
            
        Case Else
            prompt = prompt & "Provide a comprehensive analysis of this data with key insights, patterns, and actionable business recommendations."
    End Select
    
    BuildAnalysisPrompt = prompt
End Function

' Build question prompt
Private Function BuildQuestionPrompt(dataArray As Variant, question As String) As String
    Dim prompt As String
    Dim headers As String
    Dim j As Integer
    Dim rowCount As Integer, colCount As Integer
    
    rowCount = UBound(dataArray, 1) - 1
    colCount = UBound(dataArray, 2)
    
    ' Extract headers
    For j = 1 To colCount
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(dataArray(1, j))
    Next j
    
    prompt = "Based on this dataset with " & rowCount & " rows and " & colCount & " columns:" & vbCrLf & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "User Question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Please provide a clear, detailed, and actionable answer based on the data characteristics. "
    prompt = prompt & "Include specific insights, patterns, and business recommendations where applicable."
    
    BuildQuestionPrompt = prompt
End Function

' Write results to new Excel sheet
Private Sub WriteResultsToSheet(results As String, sheetName As String)
    Dim ws As Worksheet
    Dim resultLines As Variant
    Dim i As Integer
    
    ' Delete existing sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = Worksheets.Add
    ws.Name = sheetName
    
    ' Split results into lines and write to sheet
    resultLines = Split(results, vbCrLf)
    
    For i = 0 To UBound(resultLines)
        ws.Cells(i + 1, 1).Value = resultLines(i)
    Next i
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .AutoFit
    End With
    
    ' Set column width for better readability
    ws.Columns(1).ColumnWidth = 100
    
    ' Activate the sheet
    ws.Activate
    ws.Range("A1").Select
End Sub

' Escape JSON strings
Private Function EscapeJSON(text As String) As String
    EscapeJSON = Replace(text, "\", "\\")
    EscapeJSON = Replace(EscapeJSON, """", "\""")
    EscapeJSON = Replace(EscapeJSON, vbCrLf, "\n")
    EscapeJSON = Replace(EscapeJSON, vbCr, "\n")
    EscapeJSON = Replace(EscapeJSON, vbLf, "\n")
    EscapeJSON = Replace(EscapeJSON, vbTab, "\t")
End Function

' Extract response from JSON
Private Function ExtractResponseFromJSON(jsonText As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim result As String
    
    ' Simple JSON parsing - find "response" field
    startPos = InStr(jsonText, """response"":""") + 12
    If startPos > 12 Then
        endPos = InStr(startPos, jsonText, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, """}")
        If endPos > startPos Then
            result = Mid(jsonText, startPos, endPos - startPos)
            ' Unescape JSON
            result = Replace(result, "\n", vbCrLf)
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            result = Replace(result, "\t", vbTab)
            ExtractResponseFromJSON = result
        Else
            ExtractResponseFromJSON = "❌ Error parsing response format"
        End If
    Else
        ExtractResponseFromJSON = "❌ Error: No response found in server reply"
    End If
End Function

' Add custom ribbon (placeholder - requires ribbon XML)
Private Sub AddCustomRibbon()
    ' This would require custom ribbon XML for Excel 2007+
    ' For now, users can access functions via Alt+F8 or Developer tab
End Sub

' Show help information
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "🤖 EXCEL-OLLAMA AI PLUGIN HELP" & vbCrLf & vbCrLf
    helpText = helpText & "AVAILABLE FUNCTIONS:" & vbCrLf
    helpText = helpText & "• AnalyzeSelectedData - Statistical analysis of selected data" & vbCrLf
    helpText = helpText & "• AskQuestionAboutData - Ask questions in natural language" & vbCrLf
    helpText = helpText & "• GenerateComprehensiveReport - Multi-perspective analysis" & vbCrLf
    helpText = helpText & "• AnalyzeTrends - Trend and pattern analysis" & vbCrLf
    helpText = helpText & "• DetectPatterns - Anomaly and pattern detection" & vbCrLf & vbCrLf
    helpText = helpText & "CONFIGURATION:" & vbCrLf
    helpText = helpText & "• ConfigureOllamaServer - Set server URL and model" & vbCrLf
    helpText = helpText & "• TestOllamaConnection - Test server connectivity" & vbCrLf & vbCrLf
    helpText = helpText & "HOW TO USE:" & vbCrLf
    helpText = helpText & "1. Select your data range (including headers)" & vbCrLf
    helpText = helpText & "2. Press Alt+F8 to run macros" & vbCrLf
    helpText = helpText & "3. Choose the analysis function you want" & vbCrLf
    helpText = helpText & "4. Results will appear in a new sheet" & vbCrLf & vbCrLf
    helpText = helpText & "CURRENT CONFIGURATION:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel
    
    MsgBox helpText, vbInformation, "Ollama AI Plugin Help"
End Sub