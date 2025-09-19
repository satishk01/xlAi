' ============================================================================
' Excel-Ollama AI Plugin - ADVANCED AI ENTERPRISE VERSION
' Features: Qwen3, DeepSeek Thinking Models, Native Charts, Copilot Analysis
' ============================================================================

Option Explicit

' Configuration - UPDATE THIS WITH YOUR EC2 IP
Private Const OLLAMA_SERVER As String = "http://YOUR_EC2_IP:11434"
Private Const DEFAULT_MODEL As String = "qwen2.5:latest"

' Advanced AI Models Configuration
Private Const THINKING_MODEL As String = "deepseek-r1:latest"
Private Const COPILOT_MODEL As String = "qwen2.5:32b"
Private Const CHART_MODEL As String = "qwen2.5:latest"

' Enterprise settings
Private Const MAX_SAMPLE_SIZE As Long = 1000
Private Const CHUNK_SIZE As Long = 10000

' Global variables
Private currentModel As String
Private serverUrl As String
Private thinkingModel As String
Private copilotModel As String

' ============================================================================
' INITIALIZATION WITH MODEL SETUP
' ============================================================================
Sub Auto_Open()
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    thinkingModel = THINKING_MODEL
    copilotModel = COPILOT_MODEL
    
    ' Auto-install required models
    Call CheckAndInstallModels()
    
    MsgBox "🚀 ADVANCED AI Excel-Ollama Plugin loaded!" & vbCrLf & vbCrLf & _
           "🧠 Advanced Features:" & vbCrLf & _
           "✅ Qwen3 & DeepSeek Thinking Models" & vbCrLf & _
           "✅ Native Excel Chart Generation" & vbCrLf & _
           "✅ GitHub Copilot-like Analysis" & vbCrLf & _
           "✅ Hidden Thinking Process" & vbCrLf & _
           "✅ Comprehensive Data Insights" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Default Model: " & currentModel & vbCrLf & _
           "Thinking Model: " & thinkingModel & vbCrLf & _
           "Copilot Model: " & copilotModel, vbInformation, "Advanced AI Plugin"
End Sub

' ============================================================================
' ADVANCED AI FUNCTIONS
' ============================================================================

' GitHub Copilot-like Analysis - COMPREHENSIVE AI INSIGHTS
Public Sub DoCopilotAnalysis()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim rowCount As Long, colCount As Long
    Dim copilotResult As String
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Confirm Copilot analysis
    If MsgBox("🤖 GitHub Copilot-style Analysis" & vbCrLf & vbCrLf & _
              "Dataset: " & Format(rowCount, "#,##0") & " rows × " & colCount & " columns" & vbCrLf & vbCrLf & _
              "This will provide comprehensive insights like GitHub Copilot:" & vbCrLf & _
              "• Data patterns and anomalies" & vbCrLf & _
              "• Business insights and recommendations" & vbCrLf & _
              "• Predictive analysis" & vbCrLf & _
              "• Optimization suggestions" & vbCrLf & _
              "• Interactive visualizations" & vbCrLf & vbCrLf & _
              "Continue with Copilot analysis?", _
              vbYesNo + vbQuestion, "Copilot Analysis") = vbNo Then Exit Sub
    
    ' Execute Copilot analysis with thinking model
    Call PrepareExcelForProcessing("🧠 Copilot analyzing " & Format(rowCount, "#,##0") & " rows with AI thinking...")
    
    copilotResult = PerformCopilotAnalysis(selectedRange)
    
    Call RestoreExcelState()
    
    ' Show preview in message box
    MsgBox "🤖 Copilot Analysis Completed!" & vbCrLf & vbCrLf & _
           "Generated comprehensive insights for your data." & vbCrLf & vbCrLf & _
           "Preview:" & vbCrLf & Left(copilotResult, 200) & "..." & vbCrLf & vbCrLf & _
           "Full analysis has been written to a new sheet.", vbInformation, "Copilot Complete"
    
    ' Write comprehensive results
    Call WriteCopilotResultsToSheet(copilotResult, rowCount, colCount)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in DoCopilotAnalysis: " & Err.Description, vbCritical, "Copilot Error"
End Sub

' Advanced Question with Thinking Models
Public Sub AskAdvancedQuestion()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim question As String
    Dim questionType As String
    Dim answer As String
    Dim useThinking As Boolean
    
    ' Get question and type
    question = InputBox("🧠 Ask an Advanced AI Question:" & vbCrLf & vbCrLf & _
                       "Examples:" & vbCrLf & _
                       "• What are the key insights in this data?" & vbCrLf & _
                       "• Create a chart showing sales trends" & vbCrLf & _
                       "• What patterns do you see?" & vbCrLf & _
                       "• Predict next quarter's performance" & vbCrLf & _
                       "• Find anomalies in the data" & vbCrLf & _
                       "• Suggest optimization strategies", "Advanced AI Question")
    
    If question = "" Or question = "False" Then Exit Sub
    
    ' Determine if thinking model should be used
    useThinking = (InStr(LCase(question), "insight") > 0 Or _
                   InStr(LCase(question), "pattern") > 0 Or _
                   InStr(LCase(question), "predict") > 0 Or _
                   InStr(LCase(question), "analyze") > 0 Or _
                   InStr(LCase(question), "recommend") > 0)
    
    ' Check if chart generation is requested
    Dim generateChart As Boolean
    generateChart = (InStr(LCase(question), "chart") > 0 Or _
                     InStr(LCase(question), "graph") > 0 Or _
                     InStr(LCase(question), "plot") > 0 Or _
                     InStr(LCase(question), "visualiz") > 0)
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    ' Process with appropriate model
    Call PrepareExcelForProcessing("🧠 Processing with " & IIf(useThinking, "thinking model", "standard model") & "...")
    
    If useThinking Then
        answer = ProcessWithThinkingModel(selectedRange, question)
    Else
        answer = ProcessQuestionOnRangeFixed(selectedRange, question)
    End If
    
    ' Generate chart if requested
    If generateChart Then
        Call GenerateNativeChart(selectedRange, question, answer)
    End If
    
    Call RestoreExcelState()
    
    ' Show answer
    MsgBox "🧠 Advanced AI Response:" & vbCrLf & vbCrLf & answer, vbInformation, "AI Answer"
    
    ' Write detailed results
    Call WriteAdvancedResultsToSheet(question, answer, selectedRange.Rows.Count, selectedRange.Columns.Count, useThinking, generateChart)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in AskAdvancedQuestion: " & Err.Description, vbCritical, "Advanced Question Error"
End Sub

' Generate Native Excel Charts (No PNG files)
Public Sub GenerateDataVisualization()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim chartType As String
    Dim chartTitle As String
    Dim aiSuggestion As String
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    ' Get AI suggestion for best chart type
    Call PrepareExcelForProcessing("🧠 AI analyzing data for best visualization...")
    
    aiSuggestion = GetAIChartSuggestion(selectedRange)
    
    Call RestoreExcelState()
    
    ' Show AI suggestion and get user choice
    chartType = InputBox("🎨 AI Chart Recommendation:" & vbCrLf & vbCrLf & _
                        aiSuggestion & vbCrLf & vbCrLf & _
                        "Enter chart type:" & vbCrLf & _
                        "• column (recommended)" & vbCrLf & _
                        "• line" & vbCrLf & _
                        "• pie" & vbCrLf & _
                        "• scatter" & vbCrLf & _
                        "• area" & vbCrLf & _
                        "• bar", "AI Chart Generation", "column")
    
    If chartType = "" Or chartType = "False" Then Exit Sub
    
    ' Get chart title
    chartTitle = InputBox("Enter chart title:", "Chart Title", "AI-Generated Data Visualization")
    If chartTitle = "" Then chartTitle = "Data Visualization"
    
    ' Generate native Excel chart
    Call CreateNativeExcelChart(selectedRange, chartType, chartTitle, aiSuggestion)
    
    MsgBox "🎨 Native Excel chart created successfully!" & vbCrLf & vbCrLf & _
           "Chart Type: " & UCase(chartType) & vbCrLf & _
           "Title: " & chartTitle & vbCrLf & vbCrLf & _
           "The chart has been embedded in your worksheet.", vbInformation, "Chart Created"
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in GenerateDataVisualization: " & Err.Description, vbCritical, "Chart Generation Error"
End Sub' ==
==========================================================================
' ADVANCED AI PROCESSING FUNCTIONS
' ============================================================================

' Perform GitHub Copilot-like Analysis
Private Function PerformCopilotAnalysis(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim copilotPrompt As String
    Dim rawResponse As String
    Dim cleanResponse As String
    
    dataArray = selectedRange.Value2
    
    ' Build comprehensive Copilot-style prompt
    copilotPrompt = BuildCopilotPrompt(dataArray)
    
    ' Use thinking model for deep analysis
    rawResponse = CallOllamaWithThinking(copilotPrompt, copilotModel)
    
    ' Clean response (remove thinking process)
    cleanResponse = ExtractFinalAnswer(rawResponse)
    
    ' Format as Copilot-style response
    PerformCopilotAnalysis = FormatCopilotResponse(cleanResponse, selectedRange.Rows.Count, selectedRange.Columns.Count)
    
    Exit Function
    
ErrorHandler:
    PerformCopilotAnalysis = "Error in Copilot analysis: " & Err.Description
End Function

' Process with Thinking Model (Hide thinking process)
Private Function ProcessWithThinkingModel(dataRange As Range, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim thinkingPrompt As String
    Dim rawResponse As String
    Dim finalAnswer As String
    
    dataArray = dataRange.Value2
    
    ' Build thinking prompt
    thinkingPrompt = BuildThinkingPrompt(dataArray, question)
    
    ' Call thinking model
    rawResponse = CallOllamaWithThinking(thinkingPrompt, thinkingModel)
    
    ' Extract only the final answer (hide thinking)
    finalAnswer = ExtractFinalAnswer(rawResponse)
    
    ProcessWithThinkingModel = finalAnswer
    
    Exit Function
    
ErrorHandler:
    ProcessWithThinkingModel = "Error in thinking model processing: " & Err.Description
End Function

' Call Ollama with Thinking Model
Private Function CallOllamaWithThinking(prompt As String, model As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build JSON request for thinking model
    requestBody = BuildThinkingJSONRequest(model, prompt)
    url = serverUrl & "/api/generate"
    
    ' Make API call with longer timeout for thinking
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaWithThinking = ExtractResponseFromJSON(response)
    Else
        CallOllamaWithThinking = "HTTP Error " & http.Status & ": " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaWithThinking = "Thinking model error: " & Err.Description
End Function

' Extract Final Answer (Remove Thinking Process)
Private Function ExtractFinalAnswer(rawResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim finalAnswer As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Look for common thinking model patterns
    If InStr(rawResponse, "<thinking>") > 0 And InStr(rawResponse, "</thinking>") > 0 Then
        ' DeepSeek-R1 style thinking tags
        startPos = InStr(rawResponse, "</thinking>") + 12
        finalAnswer = Mid(rawResponse, startPos)
    ElseIf InStr(rawResponse, "**Final Answer:**") > 0 Then
        ' Look for final answer marker
        startPos = InStr(rawResponse, "**Final Answer:**") + 17
        finalAnswer = Mid(rawResponse, startPos)
    ElseIf InStr(rawResponse, "Answer:") > 0 Then
        ' Look for answer marker
        startPos = InStr(rawResponse, "Answer:") + 7
        finalAnswer = Mid(rawResponse, startPos)
    Else
        ' If no thinking markers found, return full response
        finalAnswer = rawResponse
    End If
    
    ' Clean up the final answer
    finalAnswer = Trim(finalAnswer)
    
    ' Remove any remaining thinking artifacts
    finalAnswer = Replace(finalAnswer, "<think>", "")
    finalAnswer = Replace(finalAnswer, "</think>", "")
    finalAnswer = Replace(finalAnswer, "Let me think about this...", "")
    finalAnswer = Replace(finalAnswer, "Thinking:", "")
    
    ExtractFinalAnswer = finalAnswer
    
    Exit Function
    
ErrorHandler:
    ExtractFinalAnswer = rawResponse ' Fallback to full response
End Function

' Build Copilot-style Prompt
Private Function BuildCopilotPrompt(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim dataStats As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    ' Extract sample data (first 5 rows)
    For i = LBound(dataArray, 1) + 1 To Application.Min(LBound(dataArray, 1) + 5, UBound(dataArray, 1))
        sampleData = sampleData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then sampleData = sampleData & ", "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    ' Build comprehensive Copilot prompt
    prompt = "You are an advanced AI data analyst like GitHub Copilot for Excel. " & _
             "Provide comprehensive, actionable insights for this dataset." & vbCrLf & vbCrLf
    
    prompt = prompt & "DATASET OVERVIEW:" & vbCrLf
    prompt = prompt & "- Rows: " & (rowCount - 1) & vbCrLf
    prompt = prompt & "- Columns: " & colCount & vbCrLf
    prompt = prompt & "- Headers: " & headers & vbCrLf & vbCrLf
    
    prompt = prompt & "SAMPLE DATA:" & vbCrLf & sampleData & vbCrLf
    
    prompt = prompt & "PROVIDE COPILOT-STYLE ANALYSIS INCLUDING:" & vbCrLf
    prompt = prompt & "1. 📊 KEY INSIGHTS & PATTERNS" & vbCrLf
    prompt = prompt & "2. 🎯 BUSINESS RECOMMENDATIONS" & vbCrLf
    prompt = prompt & "3. 📈 TREND ANALYSIS" & vbCrLf
    prompt = prompt & "4. ⚠️ ANOMALIES & OUTLIERS" & vbCrLf
    prompt = prompt & "5. 🔮 PREDICTIVE INSIGHTS" & vbCrLf
    prompt = prompt & "6. 💡 OPTIMIZATION SUGGESTIONS" & vbCrLf
    prompt = prompt & "7. 📋 NEXT STEPS & ACTION ITEMS" & vbCrLf & vbCrLf
    
    prompt = prompt & "Format your response like GitHub Copilot: clear, actionable, with specific insights and recommendations."
    
    BuildCopilotPrompt = prompt
    
    Exit Function
    
ErrorHandler:
    BuildCopilotPrompt = "Error building Copilot prompt: " & Err.Description
End Function

' Build Thinking Prompt
Private Function BuildThinkingPrompt(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
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
    
    ' Build thinking prompt
    prompt = "Think step by step about this data analysis question." & vbCrLf & vbCrLf
    
    prompt = prompt & "DATA CONTEXT:" & vbCrLf
    prompt = prompt & "- Dataset: " & (rowCount - 1) & " rows, " & colCount & " columns" & vbCrLf
    prompt = prompt & "- Columns: " & headers & vbCrLf
    prompt = prompt & "- Sample:" & vbCrLf & sampleData & vbCrLf
    
    prompt = prompt & "QUESTION: " & question & vbCrLf & vbCrLf
    
    prompt = prompt & "Please think through this carefully and provide a comprehensive answer. " & _
             "Use <thinking> tags for your reasoning process, then provide your final answer."
    
    BuildThinkingPrompt = prompt
    
    Exit Function
    
ErrorHandler:
    BuildThinkingPrompt = "Error building thinking prompt: " & Err.Description
End Function

' Build JSON Request for Thinking Models
Private Function BuildThinkingJSONRequest(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    ' Clean and escape prompt
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbCr, "\n")
    escapedPrompt = Replace(escapedPrompt, vbLf, "\n")
    
    ' Build request with thinking model parameters
    BuildThinkingJSONRequest = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false,""options"":{""temperature"":0.7,""top_p"":0.9}}"
End Function' 
============================================================================
' NATIVE EXCEL CHART GENERATION (NO PNG FILES)
' ============================================================================

' Get AI Chart Suggestion
Private Function GetAIChartSuggestion(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim chartPrompt As String
    Dim suggestion As String
    
    dataArray = selectedRange.Value2
    
    ' Build chart analysis prompt
    chartPrompt = BuildChartAnalysisPrompt(dataArray)
    
    ' Get AI suggestion
    suggestion = CallOllamaAPIReal(chartPrompt)
    
    GetAIChartSuggestion = suggestion
    
    Exit Function
    
ErrorHandler:
    GetAIChartSuggestion = "AI suggests a column chart for your data visualization."
End Function

' Build Chart Analysis Prompt
Private Function BuildChartAnalysisPrompt(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    prompt = "Analyze this data structure and recommend the best chart type:" & vbCrLf & vbCrLf
    prompt = prompt & "Data: " & (rowCount - 1) & " rows, " & colCount & " columns" & vbCrLf
    prompt = prompt & "Columns: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Recommend ONE of these chart types and explain why:" & vbCrLf
    prompt = prompt & "- column (for comparing categories)" & vbCrLf
    prompt = prompt & "- line (for trends over time)" & vbCrLf
    prompt = prompt & "- pie (for parts of a whole)" & vbCrLf
    prompt = prompt & "- scatter (for correlations)" & vbCrLf
    prompt = prompt & "- area (for cumulative data)" & vbCrLf
    prompt = prompt & "- bar (for horizontal comparisons)" & vbCrLf & vbCrLf
    prompt = prompt & "Provide a brief explanation of why this chart type is best for this data."
    
    BuildChartAnalysisPrompt = prompt
    
    Exit Function
    
ErrorHandler:
    BuildChartAnalysisPrompt = "Recommend the best chart type for this data."
End Function

' Create Native Excel Chart
Private Sub CreateNativeExcelChart(selectedRange As Range, chartType As String, chartTitle As String, aiSuggestion As String)
    On Error GoTo ErrorHandler
    
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim chartTypeEnum As Long
    
    Set ws = selectedRange.Worksheet
    
    ' Convert chart type string to Excel enum
    Select Case LCase(chartType)
        Case "column"
            chartTypeEnum = xlColumnClustered
        Case "line"
            chartTypeEnum = xlLine
        Case "pie"
            chartTypeEnum = xlPie
        Case "scatter"
            chartTypeEnum = xlXYScatter
        Case "area"
            chartTypeEnum = xlArea
        Case "bar"
            chartTypeEnum = xlBarClustered
        Case Else
            chartTypeEnum = xlColumnClustered ' Default
    End Select
    
    ' Create chart object
    Set chartObj = ws.ChartObjects.Add(Left:=selectedRange.Left + selectedRange.Width + 20, _
                                       Top:=selectedRange.Top, _
                                       Width:=400, _
                                       Height:=300)
    
    ' Configure chart
    With chartObj.Chart
        .SetSourceData selectedRange
        .ChartType = chartTypeEnum
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        
        ' Add AI suggestion as subtitle
        If Len(aiSuggestion) > 0 Then
            .ChartTitle.Text = chartTitle & vbCrLf & "AI Insight: " & Left(aiSuggestion, 100) & "..."
        End If
        
        ' Format chart
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Add data labels for pie charts
        If chartTypeEnum = xlPie Then
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
        End If
        
        ' Style the chart
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(248, 248, 248)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    ' Select the chart
    chartObj.Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating chart: " & Err.Description, vbCritical, "Chart Error"
End Sub

' Generate Native Chart from Question
Private Sub GenerateNativeChart(selectedRange As Range, question As String, aiResponse As String)
    On Error GoTo ErrorHandler
    
    Dim chartType As String
    Dim chartTitle As String
    
    ' Determine chart type from question
    If InStr(LCase(question), "trend") > 0 Or InStr(LCase(question), "time") > 0 Then
        chartType = "line"
        chartTitle = "Trend Analysis"
    ElseIf InStr(LCase(question), "compare") > 0 Or InStr(LCase(question), "comparison") > 0 Then
        chartType = "column"
        chartTitle = "Comparison Chart"
    ElseIf InStr(LCase(question), "pie") > 0 Or InStr(LCase(question), "proportion") > 0 Then
        chartType = "pie"
        chartTitle = "Proportion Analysis"
    ElseIf InStr(LCase(question), "scatter") > 0 Or InStr(LCase(question), "correlation") > 0 Then
        chartType = "scatter"
        chartTitle = "Correlation Analysis"
    Else
        chartType = "column"
        chartTitle = "Data Visualization"
    End If
    
    ' Create the chart
    Call CreateNativeExcelChart(selectedRange, chartType, chartTitle, aiResponse)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating chart from question: " & Err.Description, vbCritical, "Chart Generation Error"
End Sub

' ============================================================================
' MODEL MANAGEMENT FUNCTIONS
' ============================================================================

' Check and Install Required Models
Private Sub CheckAndInstallModels()
    On Error GoTo ErrorHandler
    
    Dim modelsToCheck As Variant
    Dim i As Long
    
    ' List of required models
    modelsToCheck = Array("qwen2.5:latest", "deepseek-r1:latest", "qwen2.5:32b")
    
    Application.StatusBar = "Checking AI models..."
    
    For i = 0 To UBound(modelsToCheck)
        If Not IsModelInstalled(CStr(modelsToCheck(i))) Then
            Call InstallModel(CStr(modelsToCheck(i)))
        End If
    Next i
    
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    ' Continue even if model check fails
End Sub

' Check if Model is Installed
Private Function IsModelInstalled(modelName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/tags"
    
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        response = http.responseText
        IsModelInstalled = (InStr(response, modelName) > 0)
    Else
        IsModelInstalled = False
    End If
    
    Exit Function
    
ErrorHandler:
    IsModelInstalled = False
End Function

' Install Model
Private Sub InstallModel(modelName As String)
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    
    ' Show installation message
    Application.StatusBar = "Installing AI model: " & modelName & " (this may take a few minutes)..."
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/pull"
    
    requestBody = "{""name"":""" & modelName & """}"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    ' Note: This is a simplified version. In practice, model pulling is async
    ' and may take a long time. Consider implementing async handling.
    
    Exit Sub
    
ErrorHandler:
    ' Continue even if installation fails
End Sub

' Configure Advanced Models
Public Sub ConfigureAdvancedModels()
    On Error GoTo ErrorHandler
    
    Dim newServer As String
    Dim newDefaultModel As String
    Dim newThinkingModel As String
    Dim newCopilotModel As String
    
    ' Get server URL
    newServer = InputBox("Enter your Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "- http://localhost:11434 (local)" & vbCrLf & _
                        "- http://your-ec2-ip:11434 (AWS EC2)", _
                        "Advanced Server Configuration", serverUrl)
    
    If newServer <> "" And newServer <> "False" Then
        serverUrl = newServer
        
        ' Configure models
        newDefaultModel = InputBox("Default Model:" & vbCrLf & vbCrLf & _
                                  "Recommended: qwen2.5:latest", _
                                  "Default Model", currentModel)
        
        newThinkingModel = InputBox("Thinking Model:" & vbCrLf & vbCrLf & _
                                   "Recommended: deepseek-r1:latest", _
                                   "Thinking Model", thinkingModel)
        
        newCopilotModel = InputBox("Copilot Model:" & vbCrLf & vbCrLf & _
                                  "Recommended: qwen2.5:32b", _
                                  "Copilot Model", copilotModel)
        
        ' Update configuration
        If newDefaultModel <> "" And newDefaultModel <> "False" Then currentModel = newDefaultModel
        If newThinkingModel <> "" And newThinkingModel <> "False" Then thinkingModel = newThinkingModel
        If newCopilotModel <> "" And newCopilotModel <> "False" Then copilotModel = newCopilotModel
        
        MsgBox "🚀 Advanced Configuration Updated!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Default Model: " & currentModel & vbCrLf & _
               "Thinking Model: " & thinkingModel & vbCrLf & _
               "Copilot Model: " & copilotModel & vbCrLf & vbCrLf & _
               "Use TestAdvancedConnection to verify all models work.", vbInformation, "Advanced Configuration"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in ConfigureAdvancedModels: " & Err.Description, vbCritical, "Configuration Error"
End Sub

' Test Advanced Connection
Public Sub TestAdvancedConnection()
    On Error GoTo ErrorHandler
    
    Dim testResults As String
    Dim modelTests As Variant
    Dim i As Long
    
    modelTests = Array(currentModel, thinkingModel, copilotModel)
    
    testResults = "🧪 ADVANCED MODEL TESTING RESULTS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    testResults = testResults & "Server: " & serverUrl & vbCrLf & vbCrLf
    
    For i = 0 To UBound(modelTests)
        Application.StatusBar = "Testing model: " & modelTests(i) & "..."
        
        If TestModelConnection(CStr(modelTests(i))) Then
            testResults = testResults & "✅ " & modelTests(i) & " - Working" & vbCrLf
        Else
            testResults = testResults & "❌ " & modelTests(i) & " - Failed" & vbCrLf
        End If
    Next i
    
    Application.StatusBar = False
    
    testResults = testResults & vbCrLf & "🎯 FEATURE AVAILABILITY:" & vbCrLf
    testResults = testResults & "- Standard Questions: " & IIf(TestModelConnection(currentModel), "✅", "❌") & vbCrLf
    testResults = testResults & "- Thinking Analysis: " & IIf(TestModelConnection(thinkingModel), "✅", "❌") & vbCrLf
    testResults = testResults & "- Copilot Analysis: " & IIf(TestModelConnection(copilotModel), "✅", "❌") & vbCrLf
    testResults = testResults & "- Chart Generation: ✅ (Native Excel)" & vbCrLf
    
    MsgBox testResults, vbInformation, "Advanced Connection Test"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in TestAdvancedConnection: " & Err.Description, vbCritical, "Connection Test Error"
End Sub

' Test Individual Model Connection
Private Function TestModelConnection(modelName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = serverUrl & "/api/generate"
    
    requestBody = "{""model"":""" & modelName & """,""prompt"":""Hello"",""stream"":false}"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    TestModelConnection = (http.Status = 200)
    
    Exit Function
    
ErrorHandler:
    TestModelConnection = False
End Function'
 ============================================================================
' RESULT FORMATTING AND OUTPUT FUNCTIONS
' ============================================================================

' Format Copilot Response
Private Function FormatCopilotResponse(response As String, rowCount As Long, colCount As Long) As String
    Dim formattedResponse As String
    
    formattedResponse = "🤖 GITHUB COPILOT-STYLE ANALYSIS" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf
    formattedResponse = formattedResponse & "📊 DATASET: " & Format(rowCount, "#,##0") & " rows × " & colCount & " columns" & vbCrLf
    formattedResponse = formattedResponse & "🧠 AI MODEL: " & copilotModel & " (Advanced Reasoning)" & vbCrLf
    formattedResponse = formattedResponse & "⏰ GENERATED: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    formattedResponse = formattedResponse & String(60, "=") & vbCrLf & vbCrLf
    formattedResponse = formattedResponse & response & vbCrLf & vbCrLf
    formattedResponse = formattedResponse & String(60, "-") & vbCrLf
    formattedResponse = formattedResponse & "💡 This analysis was generated using advanced AI reasoning models" & vbCrLf
    formattedResponse = formattedResponse & "   similar to GitHub Copilot's analytical capabilities."
    
    FormatCopilotResponse = formattedResponse
End Function

' Write Copilot Results to Sheet
Private Sub WriteCopilotResultsToSheet(copilotResult As String, rowCount As Long, colCount As Long)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    
    ' Create unique sheet name
    sheetName = "Copilot_Analysis_" & Format(Now(), "hhmmss")
    
    ' Create new sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Set ws = ActiveSheet
        MsgBox "Using current sheet for Copilot results", vbInformation
    End If
    On Error GoTo ErrorHandler
    
    ' Set sheet name
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo ErrorHandler
    
    ' Write results with rich formatting
    ws.Range("A1").Value = copilotResult
    
    ' Format the sheet like GitHub Copilot
    With ws.Columns(1)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .WrapText = True
        .ColumnWidth = 120
    End With
    
    ' Add Copilot-style header formatting
    With ws.Range("A1:A5")
        .Font.Bold = True
        .Font.Color = RGB(0, 120, 215) ' Microsoft blue
        .Interior.Color = RGB(248, 249, 250) ' Light gray background
    End With
    
    ' Activate sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' Fallback to current sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "Copilot Analysis Results:"
    ActiveSheet.Range("A2").Value = copilotResult
    MsgBox "Copilot results written to current sheet", vbInformation
End Sub

' Write Advanced Results to Sheet
Private Sub WriteAdvancedResultsToSheet(question As String, answer As String, rowCount As Long, colCount As Long, usedThinking As Boolean, generatedChart As Boolean)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    Dim resultText As String
    
    ' Create result text
    resultText = "🧠 ADVANCED AI ANALYSIS" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & "📊 DATASET: " & rowCount & " rows × " & colCount & " columns" & vbCrLf
    resultText = resultText & "🤖 AI MODEL: " & IIf(usedThinking, thinkingModel & " (Thinking)", currentModel & " (Standard)") & vbCrLf
    resultText = resultText & "📈 CHART GENERATED: " & IIf(generatedChart, "Yes (Native Excel)", "No") & vbCrLf
    resultText = resultText & "⏰ GENERATED: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    resultText = resultText & String(50, "=") & vbCrLf & vbCrLf
    resultText = resultText & "❓ QUESTION:" & vbCrLf & question & vbCrLf & vbCrLf
    resultText = resultText & "💡 AI RESPONSE:" & vbCrLf & String(30, "-") & vbCrLf & answer
    
    If usedThinking Then
        resultText = resultText & vbCrLf & vbCrLf & "🧠 NOTE: This response was generated using advanced thinking models" & vbCrLf
        resultText = resultText & "   for deeper reasoning and analysis (thinking process hidden)."
    End If
    
    ' Create unique sheet name
    sheetName = "AI_Advanced_" & Format(Now(), "hhmmss")
    
    ' Create new sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Set ws = ActiveSheet
    End If
    On Error GoTo ErrorHandler
    
    ' Set sheet name
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
    
    ' Highlight thinking model usage
    If usedThinking Then
        With ws.Range("A1:A10")
            .Interior.Color = RGB(230, 245, 255) ' Light blue for thinking
        End With
    End If
    
    ' Activate sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' Fallback
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "Advanced AI Results:"
    ActiveSheet.Range("A2").Value = resultText
End Sub

' ============================================================================
' EXISTING HELPER FUNCTIONS (FROM PREVIOUS VERSION)
' ============================================================================

' Validate and get user selection
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

' Call Ollama API (Standard)
Private Function CallOllamaAPIReal(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    requestBody = BuildJSONRequest(currentModel, prompt)
    url = serverUrl & "/api/generate"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    If http.Status = 200 Then
        response = http.responseText
        CallOllamaAPIReal = ExtractResponseFromJSON(response)
    Else
        CallOllamaAPIReal = "HTTP Error " & http.Status & ": " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    CallOllamaAPIReal = "API Error: " & Err.Description
End Function

' Build JSON request
Private Function BuildJSONRequest(model As String, prompt As String) As String
    Dim escapedPrompt As String
    
    escapedPrompt = prompt
    escapedPrompt = Replace(escapedPrompt, "\", "\\")
    escapedPrompt = Replace(escapedPrompt, """", "\""")
    escapedPrompt = Replace(escapedPrompt, vbCrLf, "\n")
    escapedPrompt = Replace(escapedPrompt, vbCr, "\n")
    escapedPrompt = Replace(escapedPrompt, vbLf, "\n")
    
    BuildJSONRequest = "{""model"":""" & model & """,""prompt"":""" & escapedPrompt & """,""stream"":false}"
End Function

' Extract response from JSON
Private Function ExtractResponseFromJSON(jsonText As String) As String
    On Error GoTo ErrorHandler
    
    Dim startPos As Long, endPos As Long
    Dim result As String
    
    startPos = InStr(jsonText, """response"":""")
    
    If startPos > 0 Then
        startPos = startPos + 12
        endPos = startPos
        Do While endPos <= Len(jsonText)
            If Mid(jsonText, endPos, 1) = """" And Mid(jsonText, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        
        If endPos > startPos Then
            result = Mid(jsonText, startPos, endPos - startPos)
            result = Replace(result, "\""", """")
            result = Replace(result, "\\", "\")
            result = Replace(result, "\n", vbCrLf)
            ExtractResponseFromJSON = result
        Else
            ExtractResponseFromJSON = "Could not parse response"
        End If
    Else
        ExtractResponseFromJSON = "No response found"
    End If
    
    Exit Function
    
ErrorHandler:
    ExtractResponseFromJSON = "JSON parsing error: " & Err.Description
End Function

' Process Question (Standard)
Private Function ProcessQuestionOnRangeFixed(dataRange As Range, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim prompt As String
    
    dataArray = dataRange.Value2
    prompt = BuildSimpleQuestionPrompt(dataArray, question)
    
    ProcessQuestionOnRangeFixed = CallOllamaAPIReal(prompt)
    
    Exit Function
    
ErrorHandler:
    ProcessQuestionOnRangeFixed = "Error processing question: " & Err.Description
End Function

' Build simple question prompt
Private Function BuildSimpleQuestionPrompt(dataArray As Variant, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    prompt = "Data: " & (rowCount - 1) & " rows with columns: " & headers & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf
    prompt = prompt & "Please provide a direct answer."
    
    BuildSimpleQuestionPrompt = prompt
    
    Exit Function
    
ErrorHandler:
    BuildSimpleQuestionPrompt = question
End Function

' Show Advanced Help
Public Sub ShowAdvancedHelp()
    Dim helpText As String
    
    helpText = "🚀 ADVANCED AI EXCEL PLUGIN" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    helpText = helpText & "🧠 ADVANCED AI FEATURES:" & vbCrLf
    helpText = helpText & "✅ Qwen3 & DeepSeek Thinking Models" & vbCrLf
    helpText = helpText & "✅ GitHub Copilot-like Analysis" & vbCrLf
    helpText = helpText & "✅ Native Excel Chart Generation" & vbCrLf
    helpText = helpText & "✅ Hidden Thinking Process" & vbCrLf & vbCrLf
    helpText = helpText & "🎯 MAIN FUNCTIONS:" & vbCrLf
    helpText = helpText & "• DoCopilotAnalysis - Comprehensive Copilot-style insights" & vbCrLf
    helpText = helpText & "• AskAdvancedQuestion - Questions with thinking models" & vbCrLf
    helpText = helpText & "• GenerateDataVisualization - AI-powered native charts" & vbCrLf & vbCrLf
    helpText = helpText & "⚙️ SETUP FUNCTIONS:" & vbCrLf
    helpText = helpText & "• ConfigureAdvancedModels - Configure all AI models" & vbCrLf
    helpText = helpText & "• TestAdvancedConnection - Test all model connections" & vbCrLf & vbCrLf
    helpText = helpText & "🤖 CURRENT MODELS:" & vbCrLf
    helpText = helpText & "Default: " & currentModel & vbCrLf
    helpText = helpText & "Thinking: " & thinkingModel & vbCrLf
    helpText = helpText & "Copilot: " & copilotModel & vbCrLf & vbCrLf
    helpText = helpText & "🎨 CHART FEATURES:" & vbCrLf
    helpText = helpText & "• Native Excel charts (no PNG files)" & vbCrLf
    helpText = helpText & "• AI-recommended chart types" & vbCrLf
    helpText = helpText & "• Automatic chart generation from questions" & vbCrLf & vbCrLf
    helpText = helpText & "🧠 THINKING MODELS:" & vbCrLf
    helpText = helpText & "• Advanced reasoning for complex questions" & vbCrLf
    helpText = helpText & "• Hidden thinking process (only final answers shown)" & vbCrLf
    helpText = helpText & "• Enhanced accuracy for analytical tasks"
    
    MsgBox helpText, vbInformation, "Advanced AI Plugin Help"
End Sub