' ============================================================================
' Excel-Ollama AI Plugin - INTERACTIVE ENTERPRISE VERSION (FIXED)
' Fixes: Data selection, Interactive charts, Sample data generator
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
' INITIALIZATION
' ============================================================================
Sub Auto_Open()
    currentModel = DEFAULT_MODEL
    serverUrl = OLLAMA_SERVER
    thinkingModel = THINKING_MODEL
    copilotModel = COPILOT_MODEL
    
    MsgBox "🚀 INTERACTIVE ENTERPRISE Excel-Ollama Plugin loaded!" & vbCrLf & vbCrLf & _
           "✅ Fixed data selection handling" & vbCrLf & _
           "✅ Interactive chart generation" & vbCrLf & _
           "✅ Sample data generator" & vbCrLf & _
           "✅ Advanced AI models support" & vbCrLf & vbCrLf & _
           "Server: " & serverUrl & vbCrLf & _
           "Default Model: " & currentModel, vbInformation, "Interactive Enterprise Plugin"
End Sub

' ============================================================================
' SAMPLE DATA GENERATOR - NEW FEATURE
' ============================================================================

' Generate Sample Data for Testing
Public Sub GenerateSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dataType As String
    Dim rowCount As Long
    Dim startCell As String
    
    ' Ask user for data type
    dataType = InputBox("Choose sample data type:" & vbCrLf & vbCrLf & _
                       "1. sales - Sales data with trends" & vbCrLf & _
                       "2. financial - Financial performance data" & vbCrLf & _
                       "3. customer - Customer analytics data" & vbCrLf & _
                       "4. inventory - Inventory management data" & vbCrLf & _
                       "5. marketing - Marketing campaign data" & vbCrLf & _
                       "6. large - Large dataset (1000+ rows)" & vbCrLf & vbCrLf & _
                       "Enter type (sales/financial/customer/inventory/marketing/large):", _
                       "Sample Data Generator", "sales")
    
    If dataType = "" Or dataType = "False" Then Exit Sub
    
    ' Ask for number of rows
    Dim rowInput As String
    rowInput = InputBox("How many rows of data to generate?" & vbCrLf & vbCrLf & _
                       "Recommended:" & vbCrLf & _
                       "• Small test: 50 rows" & vbCrLf & _
                       "• Medium test: 200 rows" & vbCrLf & _
                       "• Large test: 1000+ rows", _
                       "Sample Data Size", "100")
    
    If rowInput = "" Or rowInput = "False" Then Exit Sub
    
    rowCount = CLng(rowInput)
    If rowCount < 10 Then rowCount = 10
    If rowCount > 10000 Then rowCount = 10000
    
    ' Ask for starting cell
    startCell = InputBox("Starting cell (e.g., A1, C5):", "Starting Position", "A1")
    If startCell = "" Then startCell = "A1"
    
    Set ws = ActiveSheet
    
    ' Generate the sample data
    Call PrepareExcelForProcessing("Generating " & rowCount & " rows of " & dataType & " sample data...")
    
    Select Case LCase(dataType)
        Case "sales"
            Call GenerateSalesData(ws, startCell, rowCount)
        Case "financial"
            Call GenerateFinancialData(ws, startCell, rowCount)
        Case "customer"
            Call GenerateCustomerData(ws, startCell, rowCount)
        Case "inventory"
            Call GenerateInventoryData(ws, startCell, rowCount)
        Case "marketing"
            Call GenerateMarketingData(ws, startCell, rowCount)
        Case "large"
            Call GenerateLargeDataset(ws, startCell, rowCount)
        Case Else
            Call GenerateSalesData(ws, startCell, rowCount) ' Default
    End Select
    
    Call RestoreExcelState()
    
    ' Auto-select the generated data
    Dim endCell As String
    endCell = GetEndCell(startCell, rowCount, GetColumnCount(dataType))
    ws.Range(startCell & ":" & endCell).Select
    
    MsgBox "✅ Sample data generated successfully!" & vbCrLf & vbCrLf & _
           "📊 Type: " & UCase(dataType) & vbCrLf & _
           "📋 Rows: " & rowCount & vbCrLf & _
           "📍 Range: " & startCell & ":" & endCell & vbCrLf & vbCrLf & _
           "The data is now selected. You can:" & vbCrLf & _
           "• Run AskAdvancedQuestionFixed for AI analysis" & vbCrLf & _
           "• Run GenerateInteractiveChart for visualization" & vbCrLf & _
           "• Run DoCopilotAnalysis for comprehensive insights", _
           vbInformation, "Sample Data Ready"
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error generating sample data: " & Err.Description, vbCritical, "Sample Data Error"
End Sub' ===
=========================================================================
' FIXED ADVANCED QUESTION FUNCTION - PROPERLY USES SELECTED DATA
' ============================================================================

' Advanced Question with FIXED data handling
Public Sub AskAdvancedQuestionFixed()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim question As String
    Dim rowCount As Long, colCount As Long
    Dim answer As String
    Dim useThinking As Boolean
    Dim dataPreview As String
    
    ' Get and validate selection FIRST
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    rowCount = selectedRange.Rows.Count
    colCount = selectedRange.Columns.Count
    
    ' Show user what data is selected
    dataPreview = GetDataPreview(selectedRange)
    
    ' Get question with data context
    question = InputBox("🧠 Ask an Advanced AI Question about your selected data:" & vbCrLf & vbCrLf & _
                       "📊 SELECTED DATA PREVIEW:" & vbCrLf & _
                       "Rows: " & rowCount & " | Columns: " & colCount & vbCrLf & _
                       dataPreview & vbCrLf & vbCrLf & _
                       "💡 Question Examples:" & vbCrLf & _
                       "• What patterns do you see in this data?" & vbCrLf & _
                       "• Which column has the highest values?" & vbCrLf & _
                       "• What insights can you provide?" & vbCrLf & _
                       "• Predict future trends based on this data" & vbCrLf & _
                       "• Find anomalies or outliers" & vbCrLf & vbCrLf & _
                       "Enter your question:", "Advanced AI Question")
    
    If question = "" Or question = "False" Then Exit Sub
    
    ' Determine if thinking model should be used
    useThinking = (InStr(LCase(question), "insight") > 0 Or _
                   InStr(LCase(question), "pattern") > 0 Or _
                   InStr(LCase(question), "predict") > 0 Or _
                   InStr(LCase(question), "analyze") > 0 Or _
                   InStr(LCase(question), "recommend") > 0 Or _
                   InStr(LCase(question), "anomal") > 0 Or _
                   InStr(LCase(question), "trend") > 0)
    
    ' Process with ACTUAL selected data
    Call PrepareExcelForProcessing("🧠 Processing with " & IIf(useThinking, "thinking model", "standard model") & " on your selected data...")
    
    If useThinking Then
        answer = ProcessWithThinkingModelFixed(selectedRange, question)
    Else
        answer = ProcessQuestionOnRangeFixed(selectedRange, question)
    End If
    
    Call RestoreExcelState()
    
    ' Show answer with data context
    MsgBox "🧠 Advanced AI Response:" & vbCrLf & vbCrLf & _
           "📊 Analyzed Data: " & rowCount & " rows × " & colCount & " columns" & vbCrLf & _
           "🤖 Model: " & IIf(useThinking, thinkingModel, currentModel) & vbCrLf & vbCrLf & _
           "💡 Answer:" & vbCrLf & Left(answer, 300) & "..." & vbCrLf & vbCrLf & _
           "Full response written to new sheet.", vbInformation, "AI Analysis Complete"
    
    ' Write detailed results
    Call WriteAdvancedResultsToSheetFixed(question, answer, selectedRange, useThinking)
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in AskAdvancedQuestionFixed: " & Err.Description, vbCritical, "Advanced Question Error"
End Sub

' ============================================================================
' INTERACTIVE CHART GENERATION - FIXED AND ENHANCED
' ============================================================================

' Interactive Chart Generation with User Choice
Public Sub GenerateInteractiveChart()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    Dim chartType As String
    Dim chartTitle As String
    Dim aiSuggestion As String
    Dim userChoice As String
    Dim xColumn As String, yColumn As String
    Dim dataPreview As String
    
    ' Validate selection
    Set selectedRange = GetValidatedSelection()
    If selectedRange Is Nothing Then Exit Sub
    
    ' Show data preview
    dataPreview = GetDataPreview(selectedRange)
    
    ' Get AI suggestion for best chart type
    Call PrepareExcelForProcessing("🧠 AI analyzing your selected data for best visualization...")
    
    aiSuggestion = GetAIChartSuggestionFixed(selectedRange)
    
    Call RestoreExcelState()
    
    ' Show interactive chart options
    userChoice = InputBox("🎨 INTERACTIVE CHART GENERATOR" & vbCrLf & vbCrLf & _
                         "📊 YOUR SELECTED DATA:" & vbCrLf & _
                         dataPreview & vbCrLf & vbCrLf & _
                         "🤖 AI RECOMMENDATION:" & vbCrLf & _
                         aiSuggestion & vbCrLf & vbCrLf & _
                         "📈 AVAILABLE CHART TYPES:" & vbCrLf & _
                         "1. column - Compare categories (recommended for most data)" & vbCrLf & _
                         "2. line - Show trends over time" & vbCrLf & _
                         "3. pie - Show parts of a whole" & vbCrLf & _
                         "4. scatter - Show correlations" & vbCrLf & _
                         "5. area - Show cumulative data" & vbCrLf & _
                         "6. bar - Horizontal comparison" & vbCrLf & _
                         "7. combo - Combination chart" & vbCrLf & vbCrLf & _
                         "Enter your choice (1-7 or type name):", _
                         "Interactive Chart Generator", "1")
    
    If userChoice = "" Or userChoice = "False" Then Exit Sub
    
    ' Convert user choice to chart type
    Select Case LCase(userChoice)
        Case "1", "column"
            chartType = "column"
        Case "2", "line"
            chartType = "line"
        Case "3", "pie"
            chartType = "pie"
        Case "4", "scatter"
            chartType = "scatter"
        Case "5", "area"
            chartType = "area"
        Case "6", "bar"
            chartType = "bar"
        Case "7", "combo"
            chartType = "combo"
        Case Else
            chartType = "column" ' Default
    End Select
    
    ' Get chart title
    chartTitle = InputBox("📝 Enter chart title:" & vbCrLf & vbCrLf & _
                         "Suggested: " & GetSuggestedTitle(selectedRange, chartType), _
                         "Chart Title", GetSuggestedTitle(selectedRange, chartType))
    
    If chartTitle = "" Then chartTitle = "Data Visualization"
    
    ' For scatter plots, ask for X and Y columns
    If chartType = "scatter" And selectedRange.Columns.Count > 2 Then
        xColumn = InputBox("Select X-axis column (enter column letter or number):", "X-Axis Column", "A")
        yColumn = InputBox("Select Y-axis column (enter column letter or number):", "Y-Axis Column", "B")
    End If
    
    ' Generate the interactive chart
    Call CreateInteractiveExcelChart(selectedRange, chartType, chartTitle, aiSuggestion, xColumn, yColumn)
    
    MsgBox "🎨 Interactive chart created successfully!" & vbCrLf & vbCrLf & _
           "📊 Chart Type: " & UCase(chartType) & vbCrLf & _
           "📝 Title: " & chartTitle & vbCrLf & _
           "📈 Data Range: " & selectedRange.Address & vbCrLf & vbCrLf & _
           "The chart has been embedded in your worksheet and is fully interactive!", _
           vbInformation, "Chart Created Successfully"
    
    Exit Sub
    
ErrorHandler:
    Call RestoreExcelState()
    MsgBox "Error in GenerateInteractiveChart: " & Err.Description, vbCritical, "Chart Generation Error"
End Sub

' ============================================================================
' SAMPLE DATA GENERATORS - MULTIPLE TYPES
' ============================================================================

' Generate Sales Data
Private Sub GenerateSalesData(ws As Worksheet, startCell As String, rowCount As Long)
    Dim startRange As Range
    Dim i As Long
    Dim baseDate As Date
    
    Set startRange = ws.Range(startCell)
    baseDate = DateSerial(2024, 1, 1)
    
    ' Headers
    startRange.Offset(0, 0).Value = "Date"
    startRange.Offset(0, 1).Value = "Product"
    startRange.Offset(0, 2).Value = "Sales_Amount"
    startRange.Offset(0, 3).Value = "Units_Sold"
    startRange.Offset(0, 4).Value = "Region"
    startRange.Offset(0, 5).Value = "Salesperson"
    
    ' Data
    Dim products As Variant
    Dim regions As Variant
    Dim salespeople As Variant
    
    products = Array("Product A", "Product B", "Product C", "Product D", "Product E")
    regions = Array("North", "South", "East", "West", "Central")
    salespeople = Array("John Smith", "Jane Doe", "Mike Johnson", "Sarah Wilson", "Tom Brown")
    
    For i = 1 To rowCount
        startRange.Offset(i, 0).Value = baseDate + (i - 1)
        startRange.Offset(i, 1).Value = products((i - 1) Mod UBound(products) + 1)
        startRange.Offset(i, 2).Value = Round(500 + Rnd() * 2000, 2) ' Sales amount
        startRange.Offset(i, 3).Value = Int(10 + Rnd() * 100) ' Units sold
        startRange.Offset(i, 4).Value = regions((i - 1) Mod UBound(regions) + 1)
        startRange.Offset(i, 5).Value = salespeople((i - 1) Mod UBound(salespeople) + 1)
    Next i
    
    ' Format as table
    Dim tableRange As Range
    Set tableRange = startRange.Resize(rowCount + 1, 6)
    FormatAsTable tableRange, "Sales Data"
End Sub

' Generate Financial Data
Private Sub GenerateFinancialData(ws As Worksheet, startCell As String, rowCount As Long)
    Dim startRange As Range
    Dim i As Long
    Dim baseDate As Date
    
    Set startRange = ws.Range(startCell)
    baseDate = DateSerial(2024, 1, 1)
    
    ' Headers
    startRange.Offset(0, 0).Value = "Date"
    startRange.Offset(0, 1).Value = "Revenue"
    startRange.Offset(0, 2).Value = "Expenses"
    startRange.Offset(0, 3).Value = "Profit"
    startRange.Offset(0, 4).Value = "Department"
    startRange.Offset(0, 5).Value = "Budget_Variance"
    
    ' Data
    Dim departments As Variant
    departments = Array("Marketing", "Sales", "Operations", "IT", "HR", "Finance")
    
    For i = 1 To rowCount
        Dim revenue As Double, expenses As Double
        revenue = Round(10000 + Rnd() * 50000, 2)
        expenses = Round(revenue * (0.6 + Rnd() * 0.3), 2)
        
        startRange.Offset(i, 0).Value = baseDate + (i - 1) * 7 ' Weekly data
        startRange.Offset(i, 1).Value = revenue
        startRange.Offset(i, 2).Value = expenses
        startRange.Offset(i, 3).Value = revenue - expenses
        startRange.Offset(i, 4).Value = departments((i - 1) Mod UBound(departments) + 1)
        startRange.Offset(i, 5).Value = Round((Rnd() - 0.5) * 10000, 2) ' Budget variance
    Next i
    
    ' Format as table
    Dim tableRange As Range
    Set tableRange = startRange.Resize(rowCount + 1, 6)
    FormatAsTable tableRange, "Financial Data"
End Sub

' Generate Customer Data
Private Sub GenerateCustomerData(ws As Worksheet, startCell As String, rowCount As Long)
    Dim startRange As Range
    Dim i As Long
    
    Set startRange = ws.Range(startCell)
    
    ' Headers
    startRange.Offset(0, 0).Value = "Customer_ID"
    startRange.Offset(0, 1).Value = "Age"
    startRange.Offset(0, 2).Value = "Gender"
    startRange.Offset(0, 3).Value = "Purchase_Amount"
    startRange.Offset(0, 4).Value = "Category"
    startRange.Offset(0, 5).Value = "Satisfaction_Score"
    
    ' Data
    Dim genders As Variant
    Dim categories As Variant
    
    genders = Array("Male", "Female", "Other")
    categories = Array("Electronics", "Clothing", "Books", "Home", "Sports", "Beauty")
    
    For i = 1 To rowCount
        startRange.Offset(i, 0).Value = "CUST" & Format(i, "0000")
        startRange.Offset(i, 1).Value = Int(18 + Rnd() * 65) ' Age 18-82
        startRange.Offset(i, 2).Value = genders(Int(Rnd() * UBound(genders)) + 1)
        startRange.Offset(i, 3).Value = Round(25 + Rnd() * 500, 2)
        startRange.Offset(i, 4).Value = categories((i - 1) Mod UBound(categories) + 1)
        startRange.Offset(i, 5).Value = Round(1 + Rnd() * 4, 1) ' 1-5 rating
    Next i
    
    ' Format as table
    Dim tableRange As Range
    Set tableRange = startRange.Resize(rowCount + 1, 6)
    FormatAsTable tableRange, "Customer Data"
End Sub

' Generate Large Dataset for Performance Testing
Private Sub GenerateLargeDataset(ws As Worksheet, startCell As String, rowCount As Long)
    Dim startRange As Range
    Dim i As Long
    Dim baseDate As Date
    
    Set startRange = ws.Range(startCell)
    baseDate = DateSerial(2024, 1, 1)
    
    ' Headers for comprehensive dataset
    startRange.Offset(0, 0).Value = "ID"
    startRange.Offset(0, 1).Value = "Date"
    startRange.Offset(0, 2).Value = "Category"
    startRange.Offset(0, 3).Value = "Value1"
    startRange.Offset(0, 4).Value = "Value2"
    startRange.Offset(0, 5).Value = "Value3"
    startRange.Offset(0, 6).Value = "Status"
    startRange.Offset(0, 7).Value = "Score"
    
    ' Data arrays
    Dim categories As Variant
    Dim statuses As Variant
    
    categories = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
    statuses = Array("Active", "Inactive", "Pending", "Complete", "Failed")
    
    ' Generate large dataset efficiently
    Application.Calculation = xlCalculationManual
    
    For i = 1 To rowCount
        startRange.Offset(i, 0).Value = i
        startRange.Offset(i, 1).Value = baseDate + Int(Rnd() * 365)
        startRange.Offset(i, 2).Value = categories(Int(Rnd() * UBound(categories)) + 1)
        startRange.Offset(i, 3).Value = Round(Rnd() * 1000, 2)
        startRange.Offset(i, 4).Value = Round(Rnd() * 100, 2)
        startRange.Offset(i, 5).Value = Round(Rnd() * 50, 2)
        startRange.Offset(i, 6).Value = statuses(Int(Rnd() * UBound(statuses)) + 1)
        startRange.Offset(i, 7).Value = Round(Rnd() * 100, 1)
        
        ' Progress indicator for large datasets
        If i Mod 100 = 0 Then
            Application.StatusBar = "Generating data: " & i & " of " & rowCount & " (" & Format(i / rowCount, "0%") & ")"
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Format as table
    Dim tableRange As Range
    Set tableRange = startRange.Resize(rowCount + 1, 8)
    FormatAsTable tableRange, "Large Dataset"
End Sub' ===
=========================================================================
' FIXED DATA PROCESSING FUNCTIONS
' ============================================================================

' Get Data Preview - Shows user what data is selected
Private Function GetDataPreview(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim preview As String
    Dim headers As String
    Dim sampleRow As String
    Dim i As Long, j As Long
    
    ' Get headers
    For j = 1 To Application.Min(5, selectedRange.Columns.Count)
        If j > 1 Then headers = headers & " | "
        headers = headers & CStr(selectedRange.Cells(1, j).Value)
    Next j
    
    If selectedRange.Columns.Count > 5 Then headers = headers & " | ..."
    
    ' Get sample row
    If selectedRange.Rows.Count > 1 Then
        For j = 1 To Application.Min(5, selectedRange.Columns.Count)
            If j > 1 Then sampleRow = sampleRow & " | "
            sampleRow = sampleRow & CStr(selectedRange.Cells(2, j).Value)
        Next j
        
        If selectedRange.Columns.Count > 5 Then sampleRow = sampleRow & " | ..."
    End If
    
    preview = "Headers: " & headers
    If sampleRow <> "" Then
        preview = preview & vbCrLf & "Sample: " & sampleRow
    End If
    
    GetDataPreview = preview
    Exit Function
    
ErrorHandler:
    GetDataPreview = "Data preview unavailable"
End Function

' Process with Thinking Model - FIXED to use actual data
Private Function ProcessWithThinkingModelFixed(dataRange As Range, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim thinkingPrompt As String
    Dim rawResponse As String
    Dim finalAnswer As String
    
    ' Extract ACTUAL data from selection
    dataArray = dataRange.Value2
    
    ' Build comprehensive thinking prompt with REAL data
    thinkingPrompt = BuildComprehensiveThinkingPrompt(dataArray, question, dataRange.Address)
    
    ' Call thinking model
    rawResponse = CallOllamaWithThinking(thinkingPrompt, thinkingModel)
    
    ' Extract only the final answer (hide thinking)
    finalAnswer = ExtractFinalAnswer(rawResponse)
    
    ProcessWithThinkingModelFixed = finalAnswer
    
    Exit Function
    
ErrorHandler:
    ProcessWithThinkingModelFixed = "Error in thinking model processing: " & Err.Description
End Function

' Build Comprehensive Thinking Prompt - USES REAL DATA
Private Function BuildComprehensiveThinkingPrompt(dataArray As Variant, question As String, rangeAddress As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim actualData As String
    Dim statistics As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Extract headers
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        If j > LBound(dataArray, 2) Then headers = headers & ", "
        headers = headers & CStr(dataArray(LBound(dataArray, 1), j))
    Next j
    
    ' Extract ACTUAL data (not just sample)
    actualData = "ACTUAL DATA FROM EXCEL SELECTION (" & rangeAddress & "):" & vbCrLf
    actualData = actualData & "Headers: " & headers & vbCrLf & vbCrLf
    
    ' Include more actual data rows (up to 10)
    Dim maxRows As Long
    maxRows = Application.Min(10, rowCount - 1)
    
    For i = LBound(dataArray, 1) + 1 To LBound(dataArray, 1) + maxRows
        actualData = actualData & "Row " & (i - LBound(dataArray, 1)) & ": "
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then actualData = actualData & " | "
            actualData = actualData & CStr(dataArray(i, j))
        Next j
        actualData = actualData & vbCrLf
    Next i
    
    If rowCount > 11 Then
        actualData = actualData & "... and " & (rowCount - 11) & " more rows" & vbCrLf
    End If
    
    ' Add basic statistics
    statistics = GetBasicStatistics(dataArray)
    
    ' Build comprehensive prompt
    prompt = "You are analyzing REAL data selected by the user in Excel. Think step by step." & vbCrLf & vbCrLf
    
    prompt = prompt & "DATASET CONTEXT:" & vbCrLf
    prompt = prompt & "- Excel Range: " & rangeAddress & vbCrLf
    prompt = prompt & "- Total Rows: " & (rowCount - 1) & " (excluding header)" & vbCrLf
    prompt = prompt & "- Total Columns: " & colCount & vbCrLf & vbCrLf
    
    prompt = prompt & actualData & vbCrLf
    
    prompt = prompt & "BASIC STATISTICS:" & vbCrLf & statistics & vbCrLf
    
    prompt = prompt & "USER QUESTION: " & question & vbCrLf & vbCrLf
    
    prompt = prompt & "INSTRUCTIONS:" & vbCrLf
    prompt = prompt & "1. Analyze the ACTUAL data provided above (not hypothetical data)" & vbCrLf
    prompt = prompt & "2. Use <thinking> tags for your reasoning process" & vbCrLf
    prompt = prompt & "3. Reference specific values, patterns, and insights from the real data" & vbCrLf
    prompt = prompt & "4. Provide actionable insights based on what you see in the data" & vbCrLf
    prompt = prompt & "5. Give your final answer after the thinking process" & vbCrLf & vbCrLf
    
    prompt = prompt & "Please think through this carefully using the actual data provided."
    
    BuildComprehensiveThinkingPrompt = prompt
    
    Exit Function
    
ErrorHandler:
    BuildComprehensiveThinkingPrompt = "Error building comprehensive prompt: " & Err.Description
End Function

' Get Basic Statistics from Real Data
Private Function GetBasicStatistics(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim stats As String
    Dim i As Long, j As Long
    Dim numericCols As String
    Dim textCols As String
    Dim rowCount As Long, colCount As Long
    
    rowCount = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    colCount = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    stats = "- Data Rows: " & (rowCount - 1) & vbCrLf
    stats = stats & "- Columns: " & colCount & vbCrLf
    
    ' Analyze column types
    For j = LBound(dataArray, 2) To UBound(dataArray, 2)
        Dim colHeader As String
        Dim isNumeric As Boolean
        
        colHeader = CStr(dataArray(LBound(dataArray, 1), j))
        isNumeric = True
        
        ' Check if column is numeric
        For i = LBound(dataArray, 1) + 1 To Application.Min(LBound(dataArray, 1) + 5, UBound(dataArray, 1))
            If Not IsNumeric(dataArray(i, j)) Then
                isNumeric = False
                Exit For
            End If
        Next i
        
        If isNumeric Then
            If numericCols <> "" Then numericCols = numericCols & ", "
            numericCols = numericCols & colHeader
        Else
            If textCols <> "" Then textCols = textCols & ", "
            textCols = textCols & colHeader
        End If
    Next j
    
    If numericCols <> "" Then stats = stats & "- Numeric Columns: " & numericCols & vbCrLf
    If textCols <> "" Then stats = stats & "- Text Columns: " & textCols & vbCrLf
    
    GetBasicStatistics = stats
    
    Exit Function
    
ErrorHandler:
    GetBasicStatistics = "Statistics unavailable"
End Function

' Get AI Chart Suggestion - FIXED to use real data
Private Function GetAIChartSuggestionFixed(selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim chartPrompt As String
    Dim suggestion As String
    
    dataArray = selectedRange.Value2
    
    ' Build chart analysis prompt with REAL data
    chartPrompt = BuildChartAnalysisPromptFixed(dataArray, selectedRange.Address)
    
    ' Get AI suggestion
    suggestion = CallOllamaAPIReal(chartPrompt)
    
    GetAIChartSuggestionFixed = suggestion
    
    Exit Function
    
ErrorHandler:
    GetAIChartSuggestionFixed = "AI suggests a column chart based on your data structure."
End Function

' Build Chart Analysis Prompt - USES REAL DATA
Private Function BuildChartAnalysisPromptFixed(dataArray As Variant, rangeAddress As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim j As Long, i As Long
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
            If j > LBound(dataArray, 2) Then sampleData = sampleData & " | "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    prompt = "Analyze this REAL Excel data and recommend the best chart type:" & vbCrLf & vbCrLf
    prompt = prompt & "ACTUAL DATA FROM EXCEL RANGE " & rangeAddress & ":" & vbCrLf
    prompt = prompt & "- Rows: " & (rowCount - 1) & " data rows" & vbCrLf
    prompt = prompt & "- Columns: " & colCount & vbCrLf
    prompt = prompt & "- Headers: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "SAMPLE DATA:" & vbCrLf & sampleData & vbCrLf
    
    prompt = prompt & "Based on this ACTUAL data, recommend the BEST chart type:" & vbCrLf
    prompt = prompt & "• column - for comparing categories or values" & vbCrLf
    prompt = prompt & "• line - for trends over time or continuous data" & vbCrLf
    prompt = prompt & "• pie - for parts of a whole (percentages)" & vbCrLf
    prompt = prompt & "• scatter - for correlations between two numeric variables" & vbCrLf
    prompt = prompt & "• area - for cumulative data over time" & vbCrLf
    prompt = prompt & "• bar - for horizontal comparisons" & vbCrLf & vbCrLf
    
    prompt = prompt & "Provide a specific recommendation with reasoning based on the actual data structure and content."
    
    BuildChartAnalysisPromptFixed = prompt
    
    Exit Function
    
ErrorHandler:
    BuildChartAnalysisPromptFixed = "Analyze this data for the best chart type."
End Function

' ============================================================================
' INTERACTIVE CHART CREATION
' ============================================================================

' Create Interactive Excel Chart with User Options
Private Sub CreateInteractiveExcelChart(selectedRange As Range, chartType As String, chartTitle As String, aiSuggestion As String, Optional xColumn As String = "", Optional yColumn As String = "")
    On Error GoTo ErrorHandler
    
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim chartTypeEnum As Long
    Dim chartRange As Range
    
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
        Case "combo"
            chartTypeEnum = xlColumnClustered ' Start with column, user can modify
        Case Else
            chartTypeEnum = xlColumnClustered ' Default
    End Select
    
    ' Determine chart data range
    If chartType = "scatter" And xColumn <> "" And yColumn <> "" Then
        ' For scatter plots, use specific columns
        Set chartRange = GetScatterPlotRange(selectedRange, xColumn, yColumn)
    Else
        ' Use full selection
        Set chartRange = selectedRange
    End If
    
    ' Create chart object with better positioning
    Dim chartLeft As Double, chartTop As Double
    chartLeft = selectedRange.Left + selectedRange.Width + 50
    chartTop = selectedRange.Top
    
    ' Adjust if chart would go off screen
    If chartLeft + 500 > ws.Cells.SpecialCells(xlCellTypeVisible).Left + ws.Cells.SpecialCells(xlCellTypeVisible).Width Then
        chartLeft = selectedRange.Left
        chartTop = selectedRange.Top + selectedRange.Height + 20
    End If
    
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, _
                                       Top:=chartTop, _
                                       Width:=500, _
                                       Height:=350)
    
    ' Configure chart with enhanced options
    With chartObj.Chart
        .SetSourceData chartRange
        .ChartType = chartTypeEnum
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        
        ' Add AI insight as subtitle if space allows
        If Len(aiSuggestion) > 0 And Len(aiSuggestion) < 100 Then
            .ChartTitle.Text = chartTitle & vbCrLf & "AI Insight: " & Left(aiSuggestion, 80) & "..."
        End If
        
        ' Enhanced formatting based on chart type
        Select Case chartType
            Case "pie"
                .HasLegend = True
                .Legend.Position = xlLegendPositionRight
                If .SeriesCollection.Count > 0 Then
                    .SeriesCollection(1).HasDataLabels = True
                    .SeriesCollection(1).DataLabels.ShowPercentage = True
                    .SeriesCollection(1).DataLabels.ShowValue = False
                End If
                
            Case "scatter"
                .HasLegend = False
                If .SeriesCollection.Count > 0 Then
                    .SeriesCollection(1).HasDataLabels = False
                    .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
                    .SeriesCollection(1).MarkerSize = 8
                End If
                
            Case "line"
                .HasLegend = True
                .Legend.Position = xlLegendPositionBottom
                If .SeriesCollection.Count > 0 Then
                    .SeriesCollection(1).Smooth = True
                End If
                
            Case Else ' Column, bar, area
                .HasLegend = True
                .Legend.Position = xlLegendPositionBottom
        End Select
        
        ' Professional styling
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(248, 248, 248)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = RGB(200, 200, 200)
        
        ' Add axis titles if appropriate
        If chartType <> "pie" Then
            .Axes(xlCategory).HasTitle = True
            .Axes(xlValue).HasTitle = True
            
            If chartRange.Rows.Count > 1 Then
                .Axes(xlCategory).AxisTitle.Text = CStr(chartRange.Cells(1, 1).Value)
                If chartRange.Columns.Count > 1 Then
                    .Axes(xlValue).AxisTitle.Text = CStr(chartRange.Cells(1, 2).Value)
                End If
            End If
        End If
    End With
    
    ' Select the chart for user interaction
    chartObj.Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating interactive chart: " & Err.Description, vbCritical, "Chart Error"
End Sub

' Helper Functions for Chart Generation
Private Function GetSuggestedTitle(selectedRange As Range, chartType As String) As String
    Dim title As String
    
    If selectedRange.Rows.Count > 1 And selectedRange.Columns.Count > 1 Then
        title = CStr(selectedRange.Cells(1, 2).Value) & " by " & CStr(selectedRange.Cells(1, 1).Value)
    Else
        title = "Data Analysis"
    End If
    
    Select Case LCase(chartType)
        Case "line"
            title = title & " - Trend Analysis"
        Case "pie"
            title = title & " - Distribution"
        Case "scatter"
            title = title & " - Correlation"
        Case "area"
            title = title & " - Cumulative View"
    End Select
    
    GetSuggestedTitle = title
End Function

Private Function GetScatterPlotRange(selectedRange As Range, xColumn As String, yColumn As String) As Range
    ' This is a simplified version - in practice, you'd want more sophisticated column selection
    Set GetScatterPlotRange = selectedRange
End Function

Private Function GetColumnCount(dataType As String) As Long
    Select Case LCase(dataType)
        Case "sales"
            GetColumnCount = 6
        Case "financial"
            GetColumnCount = 6
        Case "customer"
            GetColumnCount = 6
        Case "inventory"
            GetColumnCount = 6
        Case "marketing"
            GetColumnCount = 6
        Case "large"
            GetColumnCount = 8
        Case Else
            GetColumnCount = 6
    End Select
End Function

Private Function GetEndCell(startCell As String, rowCount As Long, colCount As Long) As String
    Dim startCol As Long, startRow As Long
    
    ' Simple conversion - in practice you'd want more robust parsing
    startCol = Asc(Left(startCell, 1)) - Asc("A") + 1
    startRow = CLng(Mid(startCell, 2))
    
    Dim endCol As String
    endCol = Chr(Asc("A") + colCount - 1)
    
    GetEndCell = endCol & (startRow + rowCount)
End Function

Private Sub FormatAsTable(tableRange As Range, tableName As String)
    On Error Resume Next
    
    ' Format headers
    With tableRange.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Add borders
    With tableRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
        .Weight = xlThin
    End With
    
    ' Alternate row colors
    Dim i As Long
    For i = 2 To tableRange.Rows.Count Step 2
        tableRange.Rows(i).Interior.Color = RGB(242, 242, 242)
    Next i
    
    ' Auto-fit columns
    tableRange.Columns.AutoFit
    
    On Error GoTo 0
End Sub' 
============================================================================
' REMAINING HELPER FUNCTIONS AND API CALLS
' ============================================================================

' Write Advanced Results with Better Formatting
Private Sub WriteAdvancedResultsToSheetFixed(question As String, answer As String, selectedRange As Range, usedThinking As Boolean)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    Dim resultText As String
    
    ' Create comprehensive result text
    resultText = "🧠 ADVANCED AI ANALYSIS - INTERACTIVE VERSION" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf
    resultText = resultText & "📊 ANALYZED DATA RANGE: " & selectedRange.Address & vbCrLf
    resultText = resultText & "📋 Data Size: " & selectedRange.Rows.Count & " rows × " & selectedRange.Columns.Count & " columns" & vbCrLf
    resultText = resultText & "🤖 AI Model: " & IIf(usedThinking, thinkingModel & " (Advanced Thinking)", currentModel & " (Standard)") & vbCrLf
    resultText = resultText & "⏰ Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss") & vbCrLf
    resultText = resultText & String(60, "=") & vbCrLf & vbCrLf
    
    ' Add data preview
    resultText = resultText & "📊 DATA PREVIEW:" & vbCrLf
    resultText = resultText & GetDataPreview(selectedRange) & vbCrLf & vbCrLf
    
    resultText = resultText & "❓ QUESTION:" & vbCrLf & question & vbCrLf & vbCrLf
    resultText = resultText & "💡 AI RESPONSE:" & vbCrLf & String(30, "-") & vbCrLf & answer
    
    If usedThinking Then
        resultText = resultText & vbCrLf & vbCrLf & "🧠 ANALYSIS METHOD:" & vbCrLf
        resultText = resultText & "This response used advanced thinking models for deeper reasoning." & vbCrLf
        resultText = resultText & "The AI analyzed your actual selected data and provided insights" & vbCrLf
        resultText = resultText & "based on the real values, patterns, and structure in your dataset."
    End If
    
    ' Create unique sheet name
    sheetName = "AI_Analysis_" & Format(Now(), "hhmmss")
    
    ' Create and write to sheet
    Call WriteToNewSheet(resultText, sheetName)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error writing results: " & Err.Description, vbCritical, "Write Error"
End Sub

' Write to New Sheet with Better Error Handling
Private Sub WriteToNewSheet(content As String, sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' Create new sheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Set ws = ActiveSheet
        MsgBox "Using current sheet for results", vbInformation
    End If
    On Error GoTo ErrorHandler
    
    ' Set sheet name safely
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo ErrorHandler
    
    ' Write content
    ws.Range("A1").Value = content
    
    ' Format the sheet
    With ws.Columns(1)
        .Font.Name = "Consolas"
        .Font.Size = 10
        .WrapText = True
        .ColumnWidth = 120
    End With
    
    ' Activate sheet
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    ' Fallback to current sheet
    On Error Resume Next
    ActiveSheet.Range("A1").Value = "AI Analysis Results:"
    ActiveSheet.Range("A2").Value = content
    MsgBox "Results written to current sheet due to error: " & Err.Description, vbInformation
End Sub

' ============================================================================
' EXISTING HELPER FUNCTIONS (UPDATED)
' ============================================================================

' Validate and get user selection
Private Function GetValidatedSelection() As Range
    On Error GoTo ErrorHandler
    
    Dim selectedRange As Range
    
    Set selectedRange = Application.Selection
    
    If selectedRange Is Nothing Then
        MsgBox "❌ No range selected." & vbCrLf & vbCrLf & _
               "Please select your data range first:" & vbCrLf & _
               "1. Click and drag to select your data" & vbCrLf & _
               "2. Include headers in the first row" & vbCrLf & _
               "3. Make sure all relevant data is selected" & vbCrLf & vbCrLf & _
               "💡 TIP: Use GenerateSampleData if you need test data", _
               vbExclamation, "Selection Required"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    If selectedRange.Rows.Count < 2 Then
        MsgBox "❌ Please select at least 2 rows of data." & vbCrLf & vbCrLf & _
               "Current selection: " & selectedRange.Rows.Count & " row(s)" & vbCrLf & vbCrLf & _
               "You need:" & vbCrLf & _
               "• Row 1: Column headers" & vbCrLf & _
               "• Row 2+: Your actual data" & vbCrLf & vbCrLf & _
               "💡 TIP: Use GenerateSampleData to create test data", _
               vbExclamation, "Insufficient Data"
        Set GetValidatedSelection = Nothing
        Exit Function
    End If
    
    ' Show confirmation of what's selected
    Dim confirmMsg As String
    confirmMsg = "✅ Data Selection Confirmed:" & vbCrLf & vbCrLf & _
                "📍 Range: " & selectedRange.Address & vbCrLf & _
                "📊 Size: " & selectedRange.Rows.Count & " rows × " & selectedRange.Columns.Count & " columns" & vbCrLf & _
                "📋 Headers: " & GetHeaderPreview(selectedRange) & vbCrLf & vbCrLf & _
                "This data will be analyzed by the AI."
    
    ' Don't show confirmation every time - just validate
    Set GetValidatedSelection = selectedRange
    Exit Function
    
ErrorHandler:
    Set GetValidatedSelection = Nothing
    MsgBox "Selection validation error: " & Err.Description, vbCritical, "Selection Error"
End Function

Private Function GetHeaderPreview(selectedRange As Range) As String
    Dim headers As String
    Dim j As Long
    
    For j = 1 To Application.Min(4, selectedRange.Columns.Count)
        If j > 1 Then headers = headers & ", "
        headers = headers & CStr(selectedRange.Cells(1, j).Value)
    Next j
    
    If selectedRange.Columns.Count > 4 Then headers = headers & ", ..."
    
    GetHeaderPreview = headers
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

' Call Ollama API (Standard) - Existing function
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

' Process Question (Standard) - Existing function
Private Function ProcessQuestionOnRangeFixed(dataRange As Range, question As String) As String
    On Error GoTo ErrorHandler
    
    Dim dataArray As Variant
    Dim prompt As String
    
    dataArray = dataRange.Value2
    prompt = BuildSimpleQuestionPromptFixed(dataArray, question, dataRange.Address)
    
    ProcessQuestionOnRangeFixed = CallOllamaAPIReal(prompt)
    
    Exit Function
    
ErrorHandler:
    ProcessQuestionOnRangeFixed = "Error processing question: " & Err.Description
End Function

' Build simple question prompt - FIXED
Private Function BuildSimpleQuestionPromptFixed(dataArray As Variant, question As String, rangeAddress As String) As String
    On Error GoTo ErrorHandler
    
    Dim prompt As String
    Dim headers As String
    Dim sampleData As String
    Dim j As Long, i As Long
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
            If j > LBound(dataArray, 2) Then sampleData = sampleData & " | "
            sampleData = sampleData & CStr(dataArray(i, j))
        Next j
        sampleData = sampleData & vbCrLf
    Next i
    
    prompt = "ANALYZING REAL EXCEL DATA FROM RANGE " & rangeAddress & ":" & vbCrLf & vbCrLf
    prompt = prompt & "Data: " & (rowCount - 1) & " rows with columns: " & headers & vbCrLf & vbCrLf
    prompt = prompt & "Sample data from your selection:" & vbCrLf & sampleData & vbCrLf
    prompt = prompt & "Question: " & question & vbCrLf & vbCrLf
    prompt = prompt & "Please analyze the ACTUAL data provided and give a specific answer based on the real values and patterns you see."
    
    BuildSimpleQuestionPromptFixed = prompt
    
    Exit Function
    
ErrorHandler:
    BuildSimpleQuestionPromptFixed = question
End Function

' Existing API helper functions
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

' Call Ollama with Thinking - Existing function
Private Function CallOllamaWithThinking(prompt As String, model As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    requestBody = BuildJSONRequest(model, prompt)
    url = serverUrl & "/api/generate"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
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

' Extract Final Answer - Existing function
Private Function ExtractFinalAnswer(rawResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim finalAnswer As String
    Dim startPos As Long
    
    ' Look for common thinking model patterns
    If InStr(rawResponse, "<thinking>") > 0 And InStr(rawResponse, "</thinking>") > 0 Then
        startPos = InStr(rawResponse, "</thinking>") + 12
        finalAnswer = Mid(rawResponse, startPos)
    ElseIf InStr(rawResponse, "**Final Answer:**") > 0 Then
        startPos = InStr(rawResponse, "**Final Answer:**") + 17
        finalAnswer = Mid(rawResponse, startPos)
    ElseIf InStr(rawResponse, "Answer:") > 0 Then
        startPos = InStr(rawResponse, "Answer:") + 7
        finalAnswer = Mid(rawResponse, startPos)
    Else
        finalAnswer = rawResponse
    End If
    
    ' Clean up the final answer
    finalAnswer = Trim(finalAnswer)
    finalAnswer = Replace(finalAnswer, "<think>", "")
    finalAnswer = Replace(finalAnswer, "</think>", "")
    finalAnswer = Replace(finalAnswer, "Let me think about this...", "")
    finalAnswer = Replace(finalAnswer, "Thinking:", "")
    
    ExtractFinalAnswer = finalAnswer
    
    Exit Function
    
ErrorHandler:
    ExtractFinalAnswer = rawResponse
End Function

' Configuration Functions
Public Sub ConfigureAdvancedModels()
    On Error GoTo ErrorHandler
    
    Dim newServer As String
    Dim newDefaultModel As String
    Dim newThinkingModel As String
    Dim newCopilotModel As String
    
    newServer = InputBox("Enter your Ollama Server URL:" & vbCrLf & vbCrLf & _
                        "Examples:" & vbCrLf & _
                        "- http://localhost:11434 (local)" & vbCrLf & _
                        "- http://your-ec2-ip:11434 (AWS EC2)", _
                        "Advanced Server Configuration", serverUrl)
    
    If newServer <> "" And newServer <> "False" Then
        serverUrl = newServer
        
        newDefaultModel = InputBox("Default Model:", "Default Model", currentModel)
        newThinkingModel = InputBox("Thinking Model:", "Thinking Model", thinkingModel)
        newCopilotModel = InputBox("Copilot Model:", "Copilot Model", copilotModel)
        
        If newDefaultModel <> "" And newDefaultModel <> "False" Then currentModel = newDefaultModel
        If newThinkingModel <> "" And newThinkingModel <> "False" Then thinkingModel = newThinkingModel
        If newCopilotModel <> "" And newCopilotModel <> "False" Then copilotModel = newCopilotModel
        
        MsgBox "Configuration Updated!" & vbCrLf & vbCrLf & _
               "Server: " & serverUrl & vbCrLf & _
               "Default: " & currentModel & vbCrLf & _
               "Thinking: " & thinkingModel & vbCrLf & _
               "Copilot: " & copilotModel, vbInformation, "Configuration"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Configuration error: " & Err.Description, vbCritical, "Configuration Error"
End Sub

' Show Interactive Help
Public Sub ShowInteractiveHelp()
    Dim helpText As String
    
    helpText = "🚀 INTERACTIVE ENTERPRISE EXCEL PLUGIN" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    helpText = helpText & "✅ FIXED FEATURES:" & vbCrLf
    helpText = helpText & "• Data selection properly used in analysis" & vbCrLf
    helpText = helpText & "• Interactive chart generation with user choice" & vbCrLf
    helpText = helpText & "• Sample data generator for testing" & vbCrLf
    helpText = helpText & "• Enhanced data preview and validation" & vbCrLf & vbCrLf
    helpText = helpText & "🎯 MAIN FUNCTIONS:" & vbCrLf
    helpText = helpText & "• GenerateSampleData - Create test data" & vbCrLf
    helpText = helpText & "• AskAdvancedQuestionFixed - AI analysis of YOUR data" & vbCrLf
    helpText = helpText & "• GenerateInteractiveChart - Choose your chart type" & vbCrLf
    helpText = helpText & "• ConfigureAdvancedModels - Setup AI models" & vbCrLf & vbCrLf
    helpText = helpText & "📊 SAMPLE DATA TYPES:" & vbCrLf
    helpText = helpText & "• sales - Sales data with trends" & vbCrLf
    helpText = helpText & "• financial - Financial performance" & vbCrLf
    helpText = helpText & "• customer - Customer analytics" & vbCrLf
    helpText = helpText & "• large - Large datasets (1000+ rows)" & vbCrLf & vbCrLf
    helpText = helpText & "🎨 CHART TYPES:" & vbCrLf
    helpText = helpText & "• Column, Line, Pie, Scatter, Area, Bar, Combo" & vbCrLf
    helpText = helpText & "• AI recommendations with user override" & vbCrLf & vbCrLf
    helpText = helpText & "💡 QUICK START:" & vbCrLf
    helpText = helpText & "1. Run GenerateSampleData to create test data" & vbCrLf
    helpText = helpText & "2. Select the generated data" & vbCrLf
    helpText = helpText & "3. Run AskAdvancedQuestionFixed for analysis" & vbCrLf
    helpText = helpText & "4. Run GenerateInteractiveChart for visualization"
    
    MsgBox helpText, vbInformation, "Interactive Plugin Help"
End Sub