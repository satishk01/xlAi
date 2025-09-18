' ============================================================================
' Excel-Ollama AI Plugin - CORRECTED VERSION
' Ensures all functions are visible in Alt+F8
' ============================================================================

Option Explicit

' Configuration - UPDATE THIS WITH YOUR EC2 IP
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const DEFAULT_MODEL As String = "llama2:latest"

' Global variables
Private currentModel As String
Private serverUrl As String

' ============================================================================
' INITIALIZATION (This runs when Excel starts)
' ============================================================================
Sub Auto_Open()
    ' This runs when the add-in loads
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    
    ' Show welcome message
    MsgBox "🤖 Excel-Ollama AI Plugin loaded!" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel & vbCrLf & vbCrLf & _
           "Press Alt+F8 to see available functions.", vbInformation, "Ollama AI Plugin"
End Sub

' ============================================================================
' PUBLIC FUNCTIONS (Visible in Alt+F8)
' ============================================================================

' 1. Analyze selected data
Public Sub AnalyzeSelectedData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range with at least 2 rows (including headers)", vbExclamation, "Ollama AI"
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

' 2. Ask question about data
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
                       "• What patterns exist?" & vbCrLf & _
                       "• Can you summarize the data?", "🤖 Ollama AI Query")
    
    If question = "" Then Exit Sub
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range first", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Processing your question..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Ask question
    answer = AskOllamaQuestion(dataArray, question)
    
    ' Write results to new sheet
    Call WriteResultsToSheet("QUESTION: " & question & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
                            "ANSWER:" & vbCrLf & answer, "AI_Query_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Query completed! Check the 'AI_Query_Results' sheet.", vbInformation, "Ollama AI"
End Sub

' 3. Generate comprehensive report
Public Sub GenerateComprehensiveReport()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim report As String
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range for report generation", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Generating comprehensive report..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Generate report
    report = GenerateMultiAnalysisReport(dataArray)
    
    ' Write results to new sheet
    Call WriteResultsToSheet(report, "AI_Comprehensive_Report")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Report generated! Check the 'AI_Comprehensive_Report' sheet.", vbInformation, "Ollama AI"
End Sub

' 4. Trend analysis
Public Sub AnalyzeTrends()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim trendResult As String
    
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range with at least 2 rows", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    Application.StatusBar = "🤖 Analyzing trends..."
    Application.ScreenUpdating = False
    
    dataArray = selectedRange.Value
    trendResult = CallOllamaAPI(dataArray, "trends")
    
    Call WriteResultsToSheet(trendResult, "AI_Trend_Analysis")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Trend analysis completed! Check the 'AI_Trend_Analysis' sheet.", vbInformation, "Ollama AI"
End Sub

' 5. Pattern detection
Public Sub DetectPatterns()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim patternResult As String
    
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range with at least 2 rows", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    Application.StatusBar = "🤖 Detecting patterns..."
    Application.ScreenUpdating = False
    
    dataArray = selectedRange.Value
    patternResult = CallOllamaAPI(dataArray, "patterns")
    
    Call WriteResultsToSheet(patternResult, "AI_Pattern_Detection")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Pattern detection completed! Check the 'AI_Pattern_Detection' sheet.", vbInformation, "Ollama AI"
End Sub

' 6. Configure server settings
Public Sub ConfigureOllamaServer()
    Dim newServer As String
    Dim newModel As String
    
    ' Get server URL
    newServer = InputBox("Enter Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "• http://localhost:11434 (local)" & vbCrLf & _
                        "• http://your-ec2-ip:11434 (EC2)", _
                        "🔧 Server Configuration", serverUrl)
    
    If newServer = "" Then Exit Sub
    
    ' Get model name
    newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                       "Available models:" & vbCrLf & _
                       "• llama2:latest (recommended)" & vbCrLf & _
                       "• mistral:latest (faster)" & vbCrLf & _
                       "• codellama:latest (code analysis)" & vbCrLf & _
                       "• phi:latest (lightweight)", _
                       "🔧 Model Configuration", currentModel)
    
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

' 7. Test connection
Public Sub TestOllamaConnection()
    Application.StatusBar = "🤖 Testing connection..."
    
    If TestConnection() Then
        MsgBox "✅ Connection successful!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel, vbInformation, "Ollama AI"
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "• Server URL: " & serverUrl & vbCrLf & _
               "• Ollama is running on EC2" & vbCrLf & _
               "• Network connectivity" & vbCrLf & _
               "• EC2 Security Group allows port 11434", vbCritical, "Ollama AI"
    End If
    
    Application.StatusBar = False
End Sub

' 8. Show help
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "🤖 EXCEL-OLLAMA AI PLUGIN HELP" & vbCrLf & vbCrLf
    helpText = helpText & "📋 AVAILABLE FUNCTIONS:" & vbCrLf
    helpText = helpText & "• AnalyzeSelectedData - Statistical analysis" & vbCrLf
    helpText = helpText & "• AskQuestionAboutData - Natural language queries" & vbCrLf
    helpText = helpText & "• GenerateComprehensiveReport - Multi-analysis report" & vbCrLf
    helpText = helpText & "• AnalyzeTrends - Trend and time series analysis" & vbCrLf
    helpText = helpText & "• DetectPatterns - Pattern and anomaly detection" & vbCrLf
    helpText = helpText & "• ConfigureOllamaServer - Change server settings" & vbCrLf
    helpText = helpText & "• TestOllamaConnection - Test connectivity" & vbCrLf & vbCrLf
    helpText = helpText & "🚀 HOW TO USE:" & vbCrLf
    helpText = helpText & "1. Select your data range (including headers)" & vbCrLf
    helpText = helpText & "2. Press Alt+F8 to open macro dialog" & vbCrLf
    helpText = helpText & "3. Choose the analysis function you want" & vbCrLf
    helpText = helpText & "4. Results appear in a new Excel sheet" & vbCrLf & vbCrLf
    helpText = helpText & "⚙️ CURRENT SETTINGS:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel & vbCrLf & vbCrLf
    helpText = helpText & "💡 TIP: Run ConfigureOllamaServer first to set your EC2 IP!"
    
    MsgBox helpText, vbInformation, "Ollama AI Plugin Help"
End Sub

' 9. Create sample data for testing
Public Sub CreateSampleData()
    Dim ws As Worksheet
    Dim sampleData As Variant
    
    ' Create new worksheet
    Set ws = Worksheets.Add
    ws.Name = "Sample_Data_" & Format(Now(), "hhmmss")
    
    ' Sample sales data
    sampleData = Array( _
        Array("Date", "Sales", "Product", "Region", "Units"), _
        Array("2024-01-01", 1000, "Product A", "North", 50), _
        Array("2024-01-02", 1200, "Product B", "South", 60), _
        Array("2024-01-03", 800, "Product A", "East", 40), _
        Array("2024-01-04", 1500, "Product C", "West", 75), _
        Array("2024-01-05", 900, "Product B", "North", 45), _
        Array("2024-01-06", 1100, "Product A", "South", 55), _
        Array("2024-01-07", 1300, "Product C", "East", 65), _
        Array("2024-01-08", 950, "Product B", "West", 48), _
        Array("2024-01-09", 1250, "Product A", "North", 62), _
        Array("2024-01-10", 1400, "Product C", "South", 70) _
    )
    
    ' Write sample data
    ws.Range("A1:E11").Value = sampleData
    
    ' Format as table
    ws.Range("A1:E11").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "SampleDataTable"
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    MsgBox "✅ Sample data created!" & vbCrLf & vbCrLf & _
           "Sheet: " & ws.Name & vbCrLf & _
           "Data: 10 rows of sales data" & vbCrLf & vbCrLf & _
           "Select the data and try AnalyzeSelectedData!", vbInformation, "Sample Data"
End Sub

' ============================================================================
' PRIVATE HELPER FUNCTIONS
' ============================================================================

' Main API call function
Private Function CallOllamaAPI(dataArray As Variant, analysisType As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim prompt As String
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build prompt
    prompt = BuildAnalysisPrompt(dataArray, analysisType)
    
    ' Build request body
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
        CallOllamaAPI = "❌ HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & vbCrLf & _
                       "Server: " & serverUrl & vbCrLf & _
                       "Please check your configuration."
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPI = "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
                   "Server: " & serverUrl & vbCrLf & _
                   "Please ensure:" & vbCrLf & _
                   "• EC2 instance is running" & vbCrLf & _
                   "• Ollama service is active" & vbCrLf & _
                   "• Security Group allows port 11434"
End Function

' Ask question function
Private Function AskOllamaQuestion(dataArray As Variant, question As String) As String
    Dim prompt As String
    
    prompt = BuildQuestionPrompt(dataArray, question)
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
        CallOllamaAPIWithPrompt = "❌ HTTP Error " & http.Status & ": " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPIWithPrompt = "❌ Connection Error: " & Err.Description
End Function

' Generate multi-analysis report
Private Function GenerateMultiAnalysisReport(dataArray As Variant) As String
    Dim report As String
    
    ' Header
    report = "🤖 COMPREHENSIVE DATA ANALYSIS REPORT" & vbCrLf
    report = report & String(60, "=") & vbCrLf
    report = report & "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    report = report & "Model: " & currentModel & vbCrLf
    report = report & "Data Size: " & (UBound(dataArray, 1) - 1) & " rows × " & UBound(dataArray, 2) & " columns" & vbCrLf
    report = report & String(60, "=") & vbCrLf & vbCrLf
    
    ' Statistical Analysis
    Application.StatusBar = "🤖 Statistical analysis..."
    report = report & "📊 STATISTICAL ANALYSIS" & vbCrLf & String(30, "-") & vbCrLf
    report = report & CallOllamaAPI(dataArray, "statistical") & vbCrLf & vbCrLf
    
    ' Trend Analysis
    Application.StatusBar = "🤖 Trend analysis..."
    report = report & "📈 TREND ANALYSIS" & vbCrLf & String(30, "-") & vbCrLf
    report = report & CallOllamaAPI(dataArray, "trends") & vbCrLf & vbCrLf
    
    ' Pattern Analysis
    Application.StatusBar = "🤖 Pattern detection..."
    report = report & "🔍 PATTERN DETECTION" & vbCrLf & String(30, "-") & vbCrLf
    report = report & CallOllamaAPI(dataArray, "patterns") & vbCrLf & vbCrLf
    
    GenerateMultiAnalysisReport = report
End Function

' Test connection
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

' Build analysis prompt
Private Function BuildAnalysisPrompt(dataArray As Variant, analysisType As String) As String
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Integer, j As Integer
    
    ' Extract headers
    For j = 1 To UBound(dataArray, 2)
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(dataArray(1, j))
    Next j
    
    ' Extract sample data (first 5 rows)
    For i = 2 To Application.Min(6, UBound(dataArray, 1))
        For j = 1 To UBound(dataArray, 2)
            If j > 1 Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build prompt
    prompt = "Analyze this dataset with " & (UBound(dataArray, 1) - 1) & " rows and " & UBound(dataArray, 2) & " columns." & vbCrLf & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Sample data:" & vbCrLf & sampleData & vbCrLf
    
    Select Case analysisType
        Case "statistical"
            prompt = prompt & "Provide comprehensive statistical analysis including summary statistics, correlations, data quality assessment, and business insights."
        Case "trends"
            prompt = prompt & "Analyze trends, patterns, seasonal variations, and provide forecasting insights with business recommendations."
        Case "patterns"
            prompt = prompt & "Detect patterns, anomalies, outliers, correlations, and clustering with business implications."
        Case Else
            prompt = prompt & "Provide comprehensive analysis with key insights and actionable recommendations."
    End Select
    
    BuildAnalysisPrompt = prompt
End Function

' Build question prompt
Private Function BuildQuestionPrompt(dataArray As Variant, question As String) As String
    Dim prompt As String
    Dim headers As String
    Dim j As Integer
    
    For j = 1 To UBound(dataArray, 2)
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(dataArray(1, j))
    Next j
    
    prompt = "Based on this dataset with " & (UBound(dataArray, 1) - 1) & " rows and " & UBound(dataArray, 2) & " columns:" & vbCrLf & vbCrLf
    prompt = prompt & "Columns: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Provide a detailed, actionable answer with specific insights and recommendations."
    
    BuildQuestionPrompt = prompt
End Function

' Write results to sheet
Private Sub WriteResultsToSheet(results As String, sheetName As String)
    Dim ws As Worksheet
    Dim resultLines As Variant
    Dim i As Integer
    
    ' Delete existing sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = Worksheets.Add
    ws.Name = sheetName
    
    ' Write results
    resultLines = Split(results, vbCrLf)
    For i = 0 To UBound(resultLines)
        ws.Cells(i + 1, 1).Value = resultLines(i)
    Next i
    
    ' Format
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 100
    End With
    
    ws.Activate
    ws.Range("A1").Select
End Sub

' JSON helpers
Private Function EscapeJSON(text As String) As String
    EscapeJSON = Replace(text, "\", "\\")
    EscapeJSON = Replace(EscapeJSON, """", "\""")
    EscapeJSON = Replace(EscapeJSON, vbCrLf, "\n")
    EscapeJSON = Replace(EscapeJSON, vbCr, "\n")
    EscapeJSON = Replace(EscapeJSON, vbLf, "\n")
End Function

Private Function ExtractResponseFromJSON(jsonText As String) As String
    Dim startPos As Integer, endPos As Integer, result As String
    
    startPos = InStr(jsonText, """response"":""") + 12
    If startPos > 12 Then
        endPos = InStr(startPos, jsonText, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, """}")
        If endPos > startPos Then
            result = Mid(jsonText, startPos, endPos - startPos)
            result = Replace(result, "\n", vbCrLf)
            result = Replace(result, "\""", """")
            ExtractResponseFromJSON = result
        Else
            ExtractResponseFromJSON = "❌ Error parsing response"
        End If
    Else
        ExtractResponseFromJSON = "❌ No response found"
    End If
End Function