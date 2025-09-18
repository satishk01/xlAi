// Excel Web Add-in for Ollama AI Plugin
// This approach uses Office.js and doesn't require Python on Windows

// Global variables
let ollamaServerUrl = 'http://localhost:11434'; // Will be configured by user
let currentModel = 'llama2:latest';

// Initialize the add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Excel-Ollama AI Plugin loaded');
        
        // Set up ribbon buttons
        setupRibbonHandlers();
        
        // Load configuration
        loadConfiguration();
    }
});

// Configuration management
function loadConfiguration() {
    // Try to load from Office settings
    Office.context.roamingSettings.get('ollamaConfig', (result) => {
        if (result.value) {
            const config = JSON.parse(result.value);
            ollamaServerUrl = config.serverUrl || ollamaServerUrl;
            currentModel = config.defaultModel || currentModel;
        }
    });
}

function saveConfiguration(config) {
    Office.context.roamingSettings.set('ollamaConfig', JSON.stringify(config));
    Office.context.roamingSettings.saveAsync();
}

// Ribbon button handlers
function setupRibbonHandlers() {
    // These functions will be called by ribbon buttons
    window.analyzeData = analyzeData;
    window.askQuestion = askQuestion;
    window.generateReport = generateReport;
    window.configurePlugin = configurePlugin;
}

// Main analysis function
async function analyzeData() {
    try {
        await Excel.run(async (context) => {
            // Get selected range
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            
            await context.sync();
            
            if (range.rowCount < 2) {
                showMessage('Please select a data range with at least 2 rows (including headers)');
                return;
            }
            
            // Show progress
            showProgress('Analyzing data...');
            
            // Convert Excel data to format suitable for analysis
            const data = processExcelData(range.values);
            
            // Send to Ollama for analysis
            const analysis = await performAnalysis(data, 'statistical_analysis');
            
            // Write results to new worksheet
            await writeResultsToSheet(context, analysis, 'AI_Analysis_Results');
            
            hideProgress();
            showMessage('Analysis completed! Check the AI_Analysis_Results sheet.');
        });
    } catch (error) {
        hideProgress();
        showError('Analysis failed: ' + error.message);
    }
}

// Natural language query function
async function askQuestion() {
    try {
        // Show input dialog
        const question = await showInputDialog('Ask a question about your data:');
        if (!question) return;
        
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (range.rowCount < 2) {
                showMessage('Please select a data range first');
                return;
            }
            
            showProgress('Processing your question...');
            
            const data = processExcelData(range.values);
            const response = await askOllamaQuestion(data, question);
            
            await writeResultsToSheet(context, response, 'AI_Query_Results');
            
            hideProgress();
            showMessage('Query completed! Check the AI_Query_Results sheet.');
        });
    } catch (error) {
        hideProgress();
        showError('Query failed: ' + error.message);
    }
}

// Generate comprehensive report
async function generateReport() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (range.rowCount < 2) {
                showMessage('Please select a data range for report generation');
                return;
            }
            
            showProgress('Generating comprehensive report...');
            
            const data = processExcelData(range.values);
            
            // Perform multiple analyses
            const analyses = await Promise.all([
                performAnalysis(data, 'statistical_analysis'),
                performAnalysis(data, 'trend_analysis'),
                performAnalysis(data, 'pattern_detection')
            ]);
            
            // Generate report
            const report = await generateComprehensiveReport(analyses);
            
            await writeResultsToSheet(context, report, 'AI_Report');
            
            hideProgress();
            showMessage('Report generated! Check the AI_Report sheet.');
        });
    } catch (error) {
        hideProgress();
        showError('Report generation failed: ' + error.message);
    }
}

// Configuration dialog
async function configurePlugin() {
    try {
        const config = await showConfigurationDialog();
        if (config) {
            ollamaServerUrl = config.serverUrl;
            currentModel = config.defaultModel;
            saveConfiguration(config);
            
            // Test connection
            const isConnected = await testOllamaConnection();
            if (isConnected) {
                showMessage('Configuration saved and connection successful!');
            } else {
                showMessage('Configuration saved but connection failed. Please check your settings.');
            }
        }
    } catch (error) {
        showError('Configuration failed: ' + error.message);
    }
}

// Data processing functions
function processExcelData(values) {
    if (!values || values.length < 2) return null;
    
    // First row as headers
    const headers = values[0];
    const rows = values.slice(1);
    
    // Convert to structured data
    const data = rows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });
    
    return {
        headers: headers,
        rows: data,
        shape: [rows.length, headers.length]
    };
}

// Ollama API functions
async function performAnalysis(data, analysisType) {
    const prompt = createAnalysisPrompt(data, analysisType);
    
    try {
        const response = await fetch(`${ollamaServerUrl}/api/generate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                model: currentModel,
                prompt: prompt,
                stream: false
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const result = await response.json();
        return {
            analysis_type: analysisType,
            response: result.response,
            model_used: currentModel,
            timestamp: new Date().toISOString()
        };
    } catch (error) {
        throw new Error(`Ollama API error: ${error.message}`);
    }
}

async function askOllamaQuestion(data, question) {
    const prompt = `
    Based on this data with ${data.shape[0]} rows and ${data.shape[1]} columns:
    
    Columns: ${data.headers.join(', ')}
    
    Sample data (first 3 rows):
    ${JSON.stringify(data.rows.slice(0, 3), null, 2)}
    
    Question: ${question}
    
    Please provide a clear, actionable answer based on the data characteristics.
    `;
    
    return await performAnalysis(data, 'custom_query');
}

async function testOllamaConnection() {
    try {
        const response = await fetch(`${ollamaServerUrl}/api/tags`, {
            method: 'GET',
            timeout: 5000
        });
        return response.ok;
    } catch (error) {
        return false;
    }
}

// Prompt creation
function createAnalysisPrompt(data, analysisType) {
    const baseInfo = `
    Analyze this dataset with ${data.shape[0]} rows and ${data.shape[1]} columns.
    
    Columns: ${data.headers.join(', ')}
    
    Sample data (first 5 rows):
    ${JSON.stringify(data.rows.slice(0, 5), null, 2)}
    `;
    
    switch (analysisType) {
        case 'statistical_analysis':
            return baseInfo + `
            Provide a statistical analysis including:
            1. Summary statistics for numeric columns
            2. Data quality assessment
            3. Key insights and patterns
            4. Recommendations for further analysis
            `;
            
        case 'trend_analysis':
            return baseInfo + `
            Analyze trends in this data:
            1. Identify time-based patterns if applicable
            2. Detect increasing/decreasing trends
            3. Seasonal patterns
            4. Forecast insights
            `;
            
        case 'pattern_detection':
            return baseInfo + `
            Detect patterns and anomalies:
            1. Unusual values or outliers
            2. Correlations between variables
            3. Clustering patterns
            4. Data quality issues
            `;
            
        default:
            return baseInfo + 'Provide a comprehensive analysis of this data.';
    }
}

// Excel worksheet functions
async function writeResultsToSheet(context, results, sheetName) {
    // Delete existing sheet if it exists
    try {
        const existingSheet = context.workbook.worksheets.getItem(sheetName);
        existingSheet.delete();
        await context.sync();
    } catch (error) {
        // Sheet doesn't exist, which is fine
    }
    
    // Create new sheet
    const sheet = context.workbook.worksheets.add(sheetName);
    
    // Write results
    const resultText = typeof results === 'string' ? results : JSON.stringify(results, null, 2);
    
    // Split into lines for better formatting
    const lines = resultText.split('\n');
    const values = lines.map(line => [line]);
    
    const range = sheet.getRange(`A1:A${lines.length}`);
    range.values = values;
    
    // Format the sheet
    range.format.font.name = 'Consolas';
    range.format.font.size = 10;
    
    // Auto-fit columns
    sheet.getUsedRange().format.autofitColumns();
    
    // Activate the sheet
    sheet.activate();
    
    await context.sync();
}

// UI Helper functions
function showMessage(message) {
    // Use Office.js notification
    Office.ribbon.requestUpdate({
        tabs: [{
            id: 'OllamaAITab',
            groups: [{
                id: 'StatusGroup',
                controls: [{
                    id: 'StatusLabel',
                    label: message
                }]
            }]
        }]
    });
    
    // Also show as dialog
    setTimeout(() => {
        alert(message);
    }, 100);
}

function showError(message) {
    console.error(message);
    alert('Error: ' + message);
}

function showProgress(message) {
    // Update status
    showMessage(message);
    
    // Show progress indicator
    document.body.style.cursor = 'wait';
}

function hideProgress() {
    document.body.style.cursor = 'default';
    showMessage('Ready');
}

async function showInputDialog(prompt) {
    return new Promise((resolve) => {
        const result = window.prompt(prompt);
        resolve(result);
    });
}

async function showConfigurationDialog() {
    return new Promise((resolve) => {
        const serverUrl = window.prompt('Enter Ollama Server URL:', ollamaServerUrl);
        if (serverUrl === null) {
            resolve(null);
            return;
        }
        
        const model = window.prompt('Enter Default Model:', currentModel);
        if (model === null) {
            resolve(null);
            return;
        }
        
        resolve({
            serverUrl: serverUrl,
            defaultModel: model
        });
    });
}

async function generateComprehensiveReport(analyses) {
    const reportPrompt = `
    Create a comprehensive business report based on these analyses:
    
    ${analyses.map((analysis, index) => `
    Analysis ${index + 1} (${analysis.analysis_type}):
    ${analysis.response}
    `).join('\n\n')}
    
    Please provide:
    1. Executive Summary
    2. Key Findings
    3. Recommendations
    4. Areas for Further Investigation
    `;
    
    try {
        const response = await fetch(`${ollamaServerUrl}/api/generate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                model: currentModel,
                prompt: reportPrompt,
                stream: false
            })
        });
        
        const result = await response.json();
        return {
            report_type: 'comprehensive_report',
            content: result.response,
            source_analyses: analyses.length,
            generated_at: new Date().toISOString()
        };
    } catch (error) {
        throw new Error(`Report generation failed: ${error.message}`);
    }
}