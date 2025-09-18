' ============================================================================
' Excel-Ollama AI Plugin - DEBUG VERSION
' Enhanced error handling and question processing
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
    
    MsgBox "🤖 Excel-Ollama AI Plugin (Debug Version) loaded!" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Model: " & currentModel & vbCrLf & vbCrLf & _
           "Enhanced error handling enabled.", vbInformation, "Ollama AI Plugin"
End Sub

' ============================================================================
' DEBUG AND TEST FUNCTIONS
' ============================================================================

' Test basic API connectivity
Public Sub TestBasicConnection()
    Dim http As Object
    Dim url As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "🔍 Testing basic connection..."
    
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        response = http.responseText
        MsgBox "✅ Connection successful!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Response length: " & Len(response) & " characters" & vbCrLf & vbCrLf & _
               "First 200 chars:" & vbCrLf & Left(response, 200), vbInformation, "Connection Test"
    Else
        MsgBox "❌ Connection failed!" & vbCrLf & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & _
               "Status Text: " & http.statusText, vbCritical, "Connection Test"
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "URL: " & url, vbCritical, "Connection Test"
End Sub

' Test simple generation
Public Sub TestSimpleGeneration()
    Dim result As String
    
    Application.StatusBar = "🔍 Testing simple generation..."
    
    result = TestOllamaGeneration("Hello, please respond with 'Test successful!'")
    
    Application.StatusBar = False
    
    MsgBox "Generation Test Result:" & vbCrLf & vbCrLf & result, vbInformation, "Generation Test"
End Sub

' Enhanced question asking with better error handling
Public Sub AskQuestionAboutDataEnhanced()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim question As String
    Dim answer As String
    Dim debugInfo As String
    
    ' Get user question
    question = InputBox("Ask a question about your data:" & vbCrLf & vbCrLf & _
                       "Examples:" & vbCrLf & _
                       "• What is the average sales?" & vbCrLf & _
                       "• Which product sells best?" & vbCrLf & _
                       "• What trends do you see?" & vbCrLf & _
                       "• Summarize this data", "🤖 Enhanced Question")
    
    If question = "" Then Exit Sub
    
    ' Check if data is selected
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range first (including headers)", vbExclamation, "Data Required"
        Exit Sub
    End If
    
    ' Show progress
    Application.StatusBar = "🤖 Processing enhanced question..."
    Application.ScreenUpdating = False
    
    ' Get data as array
    dataArray = selectedRange.Value
    
    ' Build debug info
    debugInfo = "Data Info:" & vbCrLf
    debugInfo = debugInfo & "Rows: " & UBound(dataArray, 1) & vbCrLf
    debugInfo = debugInfo & "Columns: " & UBound(dataArray, 2) & vbCrLf
    debugInfo = debugInfo & "Question: " & question & vbCrLf & vbCrLf
    
    ' Ask question with enhanced error handling
    answer = AskOllamaQuestionEnhanced(dataArray, question, debugInfo)
    
    ' Write results to new sheet
    Call WriteResultsToSheet(debugInfo & "ANSWER:" & vbCrLf & String(50, "-") & vbCrLf & answer, "AI_Enhanced_Query")
    
    ' Cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Enhanced query completed! Check the 'AI_Enhanced_Query' sheet.", vbInformation, "Query Complete"
End Sub

' ============================================================================
' ENHANCED API FUNCTIONS
' ============================================================================

' Enhanced question asking with detailed error reporting
Private Function AskOllamaQuestionEnhanced(dataArray As Variant, question As String, ByRef debugInfo As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Integer, j As Integer
    
    ' Build enhanced prompt
    prompt = BuildEnhancedQuestionPrompt(dataArray, question)
    
    ' Add to debug info
    debugInfo = debugInfo & "Prompt length: " & Len(prompt) & " characters" & vbCrLf
    debugInfo = debugInfo & "Server: " & serverUrl & vbCrLf
    debugInfo = debugInfo & "Model: " & currentModel & vbCrLf & vbCrLf
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build request body with better JSON escaping
    requestBody = BuildJSONRequest(currentModel, prompt)
    
    ' Add to debug info
    debugInfo = debugInfo & "Request body length: " & Len(requestBody) & " characters" & vbCrLf & vbCrLf
    
    ' Make API call
    url = serverUrl & "/api/generate"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.send requestBody
    
    debugInfo = debugInfo & "HTTP Status: " & http.Status & vbCrLf
    debugInfo = debugInfo & "Response length: " & Len(http.responseText) & " characters" & vbCrLf & vbCrLf
    
    If http.Status = 200 Then
        response = http.responseText
        AskOllamaQuestionEnhanced = ExtractResponseFromJSONEnhanced(response, debugInfo)
    Else
        AskOllamaQuestionEnhanced = "❌ HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & vbCrLf & _
                                   "Response: " & Left(http.responseText, 500)
    End If
    
    Exit Function
    
ErrorHandler:
    debugInfo = debugInfo & "VBA Error: " & Err.Description & " (Number: " & Err.Number & ")" & vbCrLf
    AskOllamaQuestionEnhanced = "❌ Connection Error: " & Err.Description & vbCrLf & vbCrLf & _
                               "Error Number: " & Err.Number & vbCrLf & _
                               "URL: " & url
End Function

' Test Ollama generation with simple prompt
Private Function TestOllamaGeneration(testPrompt As String) As String
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    requestBody = BuildJSONRequest(currentModel, testPrompt)
    url = serverUrl & "/api/generate"
    
    On Error GoTo ErrorHandler
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        TestOllamaGeneration = "✅ Success!" & vbCrLf & vbCrLf & _
                              "Status: " & http.Status & vbCrLf & _
                              "Response: " & ExtractResponseFromJSONSimple(response)
    Else
        TestOllamaGeneration = "❌ HTTP Error " & http.Status & ": " & http.statusText & vbCrLf & vbCrLf & _
                              "Response: " & Left(http.responseText, 300)
    End If
    
    Exit Function
    
ErrorHandler:
    TestOllamaGeneration = "❌ Error: " & Err.Description & " (Number: " & Err.Number & ")"
End Function

' Build enhanced question prompt
Private Function BuildEnhancedQuestionPrompt(dataArray As Variant, question As String) As String
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Integer, j As Integer
    Dim rowCount As Integer, colCount As Integer
    
    rowCount = UBound(dataArray, 1) - 1  ' Subtract header row
    colCount = UBound(dataArray, 2)
    
    ' Extract headers
    For j = 1 To colCount
        If j > 1 Then headers = headers & ", "
        headers = headers & """" & CStr(dataArray(1, j)) & """"
    Next j
    
    ' Extract sample data (first 3 rows only to keep prompt manageable)
    For i = 2 To Application.Min(4, UBound(dataArray, 1))
        sampleData = sampleData & "Row " & (i - 1) & ": "
        For j = 1 To colCount
            If j > 1 Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(1, j)) & "=" & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build concise prompt
    prompt = "You are analyzing a dataset with " & rowCount & " rows and " & colCount & " columns." & vbCrLf & vbCrLf
    prompt = prompt & "Column headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Sample data:" & vbCrLf & sampleData & vbCrLf
    prompt = prompt & "User question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Please provide a clear, specific answer based on the data structure and sample values shown above. "
    prompt = prompt & "Be concise and focus on what can be determined from the available information."
    
    BuildEnhancedQuestionPrompt = prompt
End Function

' Build JSON request with proper escaping
Private Function BuildJSONRequest(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    ' Enhanced JSON escaping
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbCr, "\n")
    escapedPrompt = Replace(escapedPrompt, vbLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbTab, "\t")
    escapedPrompt = Replace(escapedPrompt, Chr(8), "\b")
    escapedPrompt = Replace(escapedPrompt, Chr(12), "\f")
    
    BuildJSONRequest = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false,""options"":{""temperature"":0.7}}"
End Function

' Enhanced JSON response extraction
Private Function ExtractResponseFromJSONEnhanced(jsonText As String, ByRef debugInfo As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim result As String
    
    debugInfo = debugInfo & "JSON parsing..." & vbCrLf
    debugInfo = debugInfo & "First 200 chars of response: " & Left(jsonText, 200) & vbCrLf & vbCrLf
    
    ' Look for "response" field
    startPos = InStr(jsonText, """response"":""")
    
    If startPos > 0 Then
        startPos = startPos + 12  ' Length of "response":""
        
        ' Find the end of the response field
        endPos = startPos
        Dim inEscape As Boolean
        Dim char As String
        
        Do While endPos <= Len(jsonText)
            char = Mid(jsonText, endPos, 1)
            
            If inEscape Then
                inEscape = False
            ElseIf char = "\" Then
                inEscape = True
            ElseIf char = """" Then
                Exit Do
            End If
            
            endPos = endPos + 1
        Loop
        
        If endPos > startPos Then
            result = Mid(jsonText, startPos, endPos - startPos)
            
            ' Unescape JSON
            result = Replace(result, "\n", vbCrLf)
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            result = Replace(result, "\t", vbTab)
            
            debugInfo = debugInfo & "✅ Successfully extracted response (" & Len(result) & " characters)" & vbCrLf & vbCrLf
            ExtractResponseFromJSONEnhanced = result
        Else
            debugInfo = debugInfo & "❌ Could not find end of response field" & vbCrLf
            ExtractResponseFromJSONEnhanced = "❌ Error: Could not parse response field"
        End If
    Else
        debugInfo = debugInfo & "❌ No 'response' field found in JSON" & vbCrLf
        ExtractResponseFromJSONEnhanced = "❌ Error: No response field found in JSON" & vbCrLf & vbCrLf & _
                                         "Raw response: " & Left(jsonText, 500)
    End If
End Function

' Simple JSON response extraction for testing
Private Function ExtractResponseFromJSONSimple(jsonText As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = InStr(jsonText, """response"":""") + 12
    If startPos > 12 Then
        endPos = InStr(startPos, jsonText, """,""")
        If endPos = 0 Then endPos = InStr(startPos, jsonText, """}")
        If endPos > startPos Then
            ExtractResponseFromJSONSimple = Mid(jsonText, startPos, endPos - startPos)
        Else
            ExtractResponseFromJSONSimple = "Could not parse response"
        End If
    Else
        ExtractResponseFromJSONSimple = "No response field found"
    End If
End Function

' ============================================================================
' STANDARD FUNCTIONS (from previous version)
' ============================================================================

' Standard analyze data function
Public Sub AnalyzeSelectedData()
    Dim selectedRange As Range
    Dim dataArray As Variant
    Dim analysisResult As String
    
    Set selectedRange = Selection
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select a data range with at least 2 rows (including headers)", vbExclamation, "Ollama AI"
        Exit Sub
    End If
    
    Application.StatusBar = "🤖 Analyzing data..."
    Application.ScreenUpdating = False
    
    dataArray = selectedRange.Value
    analysisResult = CallOllamaAPIStandard(dataArray, "statistical")
    
    Call WriteResultsToSheet(analysisResult, "AI_Analysis_Results")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "✅ Analysis completed! Check the 'AI_Analysis_Results' sheet.", vbInformation, "Ollama AI"
End Sub

' Standard API call
Private Function CallOllamaAPIStandard(dataArray As Variant, analysisType As String) As String
    Dim prompt As String
    
    prompt = BuildStandardPrompt(dataArray, analysisType)
    CallOllamaAPIStandard = TestOllamaGeneration(prompt)
End Function

' Build standard analysis prompt
Private Function BuildStandardPrompt(dataArray As Variant, analysisType As String) As String
    Dim prompt As String
    Dim headers As String
    Dim j As Integer
    
    For j = 1 To UBound(dataArray, 2)
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(dataArray(1, j))
    Next j
    
    prompt = "Analyze this " & (UBound(dataArray, 1) - 1) & " row dataset with columns: " & headers & ". "
    
    Select Case analysisType
        Case "statistical"
            prompt = prompt & "Provide statistical summary including averages, trends, and key insights."
        Case Else
            prompt = prompt & "Provide comprehensive analysis with key insights."
    End Select
    
    BuildStandardPrompt = prompt
End Function

' Configuration function
Public Sub ConfigureOllamaServer()
    Dim newServer As String
    Dim newModel As String
    
    newServer = InputBox("Enter Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Example: http://your-ec2-ip:11434", _
                        "🔧 Server Configuration", serverUrl)
    
    If newServer <> "" Then
        serverUrl = newServer
        
        newModel = InputBox("Enter Model Name:" & vbCrLf & vbCrLf & _
                           "Examples: llama2:latest, mistral:latest", _
                           "🔧 Model Configuration", currentModel)
        
        If newModel <> "" Then
            currentModel = newModel
        End If
        
        MsgBox "✅ Configuration updated!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Model: " & currentModel & vbCrLf & vbCrLf & _
               "Use TestBasicConnection to verify.", vbInformation, "Configuration"
    End If
End Sub

' Write results to sheet
Private Sub WriteResultsToSheet(results As String, sheetName As String)
    Dim ws As Worksheet
    Dim resultLines As Variant
    Dim i As Integer
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add
    ws.Name = sheetName
    
    resultLines = Split(results, vbCrLf)
    For i = 0 To UBound(resultLines)
        ws.Cells(i + 1, 1).Value = resultLines(i)
    Next i
    
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 100
    End With
    
    ws.Activate
    ws.Range("A1").Select
End Sub

' Show help
Public Sub ShowHelp()
    Dim helpText As String
    
    helpText = "🤖 EXCEL-OLLAMA AI PLUGIN (DEBUG VERSION)" & vbCrLf & vbCrLf
    helpText = helpText & "🔍 DEBUG FUNCTIONS:" & vbCrLf
    helpText = helpText & "• TestBasicConnection - Test server connectivity" & vbCrLf
    helpText = helpText & "• TestSimpleGeneration - Test AI generation" & vbCrLf
    helpText = helpText & "• AskQuestionAboutDataEnhanced - Enhanced question asking" & vbCrLf & vbCrLf
    helpText = helpText & "📊 STANDARD FUNCTIONS:" & vbCrLf
    helpText = helpText & "• AnalyzeSelectedData - Basic data analysis" & vbCrLf
    helpText = helpText & "• ConfigureOllamaServer - Update server settings" & vbCrLf & vbCrLf
    helpText = helpText & "⚙️ CURRENT SETTINGS:" & vbCrLf
    helpText = helpText & "Server: " & serverUrl & vbCrLf
    helpText = helpText & "Model: " & currentModel & vbCrLf & vbCrLf
    helpText = helpText & "💡 TROUBLESHOOTING:" & vbCrLf
    helpText = helpText & "1. Run TestBasicConnection first" & vbCrLf
    helpText = helpText & "2. If that works, try TestSimpleGeneration" & vbCrLf
    helpText = helpText & "3. Then try AskQuestionAboutDataEnhanced" & vbCrLf
    helpText = helpText & "4. Check the debug output in result sheets"
    
    MsgBox helpText, vbInformation, "Debug Help"
End Sub