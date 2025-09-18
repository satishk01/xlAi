' Excel-Ollama AI Plugin - Pure VBA Solution
' No Python required on Windows!

Option Explicit

' Global configuration
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const DEFAULT_MODEL As String = "llama2:latest"

' Main analysis function
Sub AnalyzeSelectedData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    
    ' Get selected range
    Set selectedRange = Selection
    
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range with at least 2 rows (including headers)", vbExclamation
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "Analyzing data with Ollama AI..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Perform analysis
    analysisResult = CallOllamaAPI(dataArray, "statistical_analysis")
    
    ' Write results to new sheet
    Call WriteResultsToSheet(analysisResult, "AI_Analysis_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Analysis completed! Check the AI_Analysis_Results sheet.", vbInformation
End Sub

' Ask question about data
Sub AskQuestionAboutData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim question As String
    Dim answer As String
    
    ' Get user question
    question = InputBox("Ask a question about your data:", "Ollama AI Query")
    If question = "" Then Exit Sub
    
    ' Get selected range
    Set selectedRange = Selection
    
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range first", vbExclamation
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "Processing your question with Ollama AI..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Ask question
    answer = AskOllamaQuestion(dataArray, question)
    
    ' Write results to new sheet
    Call WriteResultsToSheet("Question: " & question & vbCrLf & vbCrLf & "Answer: " & answer, "AI_Query_Results")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Query completed! Check the AI_Query_Results sheet.", vbInformation
End Sub

' Generate comprehensive report
Sub GenerateAIReport()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim report As String
    
    ' Get selected range
    Set selectedRange = Selection
    
    If selectedRange.Rows.Count < 2 Then
        MsgBox "Please select a data range for report generation", vbExclamation
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "Generating comprehensive AI report..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Generate report
    report = GenerateComprehensiveReport(dataArray)
    
    ' Write results to new sheet
    Call WriteResultsToSheet(report, "AI_Comprehensive_Report")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Report generated! Check the AI_Comprehensive_Report sheet.", vbInformation
End Sub

' Core function to call Ollama API
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
    
    ' Build request body
    requestBody = "{""model"":""" & DEFAULT_MODEL & """,""prompt"":""" & EscapeJSON(prompt) & """,""stream"":false}"
    
    ' Make API call
    url = OLLAMA_SERVER & "/api/generate"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        ' Parse JSON response (simple parsing)
        response = http.responseText
        CallOllamaAPI = ExtractResponseFromJSON(response)
    Else
        CallOllamaAPI = "Error: HTTP " & http.Status & " - " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPI = "Error connecting to Ollama server: " & Err.Description
End Function

' Ask question function
Private Function AskOllamaQuestion(dataArray As Variant, question As String) As String
    Dim prompt As String
    
    ' Build question prompt
    prompt = BuildQuestionPrompt(dataArray, question)
    
    ' Call API
    AskOllamaQuestion = CallOllamaAPI(dataArray, "custom_query")
End Function

' Generate comprehensive report
Private Function GenerateComprehensiveReport(dataArray As Variant) As String
    Dim statisticalAnalysis As String
    Dim trendAnalysis As String
    Dim patternAnalysis As String
    Dim report As String
    
    ' Perform multiple analyses
    statisticalAnalysis = CallOllamaAPI(dataArray, "statistical_analysis")
    trendAnalysis = CallOllamaAPI(dataArray, "trend_analysis")
    patternAnalysis = CallOllamaAPI(dataArray, "pattern_analysis")
    
    ' Combine into comprehensive report
    report = "COMPREHENSIVE DATA ANALYSIS REPORT" & vbCrLf
    report = report & String(50, "=") & vbCrLf & vbCrLf
    report = report & "Generated: " & Now() & vbCrLf & vbCrLf
    
    report = report & "STATISTICAL ANALYSIS:" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    report = report & statisticalAnalysis & vbCrLf & vbCrLf
    
    report = report & "TREND ANALYSIS:" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    report = report & trendAnalysis & vbCrLf & vbCrLf
    
    report = report & "PATTERN ANALYSIS:" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    report = report & patternAnalysis & vbCrLf & vbCrLf
    
    GenerateComprehensiveReport = report
End Function

' Build analysis prompt
Private Function BuildAnalysisPrompt(dataArray As Variant, analysisType As String) As String
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Integer, j As Integer
    
    ' Extract headers
    For j = 1 To UBound(dataArray, 2)
        headers = headers & dataArray(1, j) & ", "
    Next j
    headers = Left(headers, Len(headers) - 2)
    
    ' Extract sample data (first 5 rows)
    For i = 2 To Application.Min(6, UBound(dataArray, 1))
        For j = 1 To UBound(dataArray, 2)
            sampleData = sampleData & dataArray(i, j) & ", "
        Next j
        sampleData = Left(sampleData, Len(sampleData) - 2) & vbCrLf
    Next i
    
    ' Build base prompt
    prompt = "Analyze this dataset with " & (UBound(dataArray, 1) - 1) & " rows and " & UBound(dataArray, 2) & " columns." & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf
    prompt = prompt & "Sample data (first 5 rows):" & vbCrLf & sampleData & vbCrLf
    
    ' Add analysis-specific instructions
    Select Case analysisType
        Case "statistical_analysis"
            prompt = prompt & "Provide a comprehensive statistical analysis including: "
            prompt = prompt & "1. Summary statistics for numeric columns "
            prompt = prompt & "2. Data quality assessment "
            prompt = prompt & "3. Key patterns and correlations "
            prompt = prompt & "4. Business insights and recommendations"
            
        Case "trend_analysis"
            prompt = prompt & "Analyze trends and patterns: "
            prompt = prompt & "1. Identify time-based trends if applicable "
            prompt = prompt & "2. Detect increasing/decreasing patterns "
            prompt = prompt & "3. Seasonal or cyclical patterns "
            prompt = prompt & "4. Forecast insights and predictions"
            
        Case "pattern_analysis"
            prompt = prompt & "Detect patterns and anomalies: "
            prompt = prompt & "1. Unusual values or outliers "
            prompt = prompt & "2. Correlations between variables "
            prompt = prompt & "3. Clustering or grouping patterns "
            prompt = prompt & "4. Data quality issues"
            
        Case Else
            prompt = prompt & "Provide a comprehensive analysis with key insights and recommendations."
    End Select
    
    BuildAnalysisPrompt = prompt
End Function

' Build question prompt
Private Function BuildQuestionPrompt(dataArray As Variant, question As String) As String
    Dim prompt As String
    Dim headers As String
    Dim j As Integer
    
    ' Extract headers
    For j = 1 To UBound(dataArray, 2)
        headers = headers & dataArray(1, j) & ", "
    Next j
    headers = Left(headers, Len(headers) - 2)
    
    prompt = "Based on this data with " & (UBound(dataArray, 1) - 1) & " rows and " & UBound(dataArray, 2) & " columns:" & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Please provide a clear, actionable answer based on the data characteristics."
    
    BuildQuestionPrompt = prompt
End Function

' Write results to new sheet
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
    
    ' Split results into lines
    resultLines = Split(results, vbCrLf)
    
    ' Write results to sheet
    For i = 0 To UBound(resultLines)
        ws.Cells(i + 1, 1).Value = resultLines(i)
    Next i
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .AutoFit
    End With
    
    ' Activate the sheet
    ws.Activate
End Sub

' Helper function to escape JSON strings
Private Function EscapeJSON(text As String) As String
    EscapeJSON = Replace(text, """", "\""")
    EscapeJSON = Replace(EscapeJSON, vbCrLf, "\n")
    EscapeJSON = Replace(EscapeJSON, vbCr, "\n")
    EscapeJSON = Replace(EscapeJSON, vbLf, "\n")
End Function

' Helper function to extract response from JSON
Private Function ExtractResponseFromJSON(jsonText As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    ' Simple JSON parsing - find "response" field
    startPos = InStr(jsonText, """response"":""") + 12
    If startPos > 12 Then
        endPos = InStr(startPos, jsonText, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, """}")
        If endPos > startPos Then
            ExtractResponseFromJSON = Mid(jsonText, startPos, endPos - startPos)
            ' Unescape JSON
            ExtractResponseFromJSON = Replace(ExtractResponseFromJSON, "\n", vbCrLf)
            ExtractResponseFromJSON = Replace(ExtractResponseFromJSON, "\""", """")
        Else
            ExtractResponseFromJSON = "Error parsing response"
        End If
    Else
        ExtractResponseFromJSON = "Error: No response found in JSON"
    End If
End Function

' Configuration function
Sub ConfigureOllamaServer()
    Dim newServer As String
    
    newServer = InputBox("Enter Ollama Server URL:", "Configuration", OLLAMA_SERVER)
    
    If newServer <> "" Then
        ' Update the constant (this would need to be done manually in the code)
        MsgBox "Please update the OLLAMA_SERVER constant in the VBA code to: " & newServer, vbInformation
    End If
End Sub

' Test connection function
Sub TestOllamaConnection()
    Dim http As Object
    Dim url As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = OLLAMA_SERVER & "/api/tags"
    
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Testing connection to Ollama server..."
    
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        MsgBox "✅ Connection successful!" & vbCrLf & "Server: " & OLLAMA_SERVER, vbInformation
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & "HTTP " & http.Status & ": " & http.statusText, vbCritical
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "❌ Connection error: " & Err.Description, vbCritical
End Sub