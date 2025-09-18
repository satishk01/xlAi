#!/bin/bash
# Web Server Setup for Excel-Ollama AI Plugin (No Python on Windows)
# This script sets up a web server on EC2 to host the Excel add-in

echo "Setting up Web-based Excel Add-in Server"
echo "========================================"

# Get EC2 public IP
EC2_PUBLIC_IP=$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)
echo "EC2 Public IP: $EC2_PUBLIC_IP"

# Update system
echo "Updating system packages..."
sudo yum update -y || sudo apt update -y

# Install Node.js and npm
echo "Installing Node.js..."
if command -v yum &> /dev/null; then
    # Amazon Linux
    curl -fsSL https://rpm.nodesource.com/setup_18.x | sudo bash -
    sudo yum install -y nodejs
elif command -v apt &> /dev/null; then
    # Ubuntu
    curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
    sudo apt-get install -y nodejs
fi

# Install Ollama (if not already installed)
echo "Installing Ollama..."
if ! command -v ollama &> /dev/null; then
    curl -fsSL https://ollama.ai/install.sh | sh
    
    # Configure Ollama for external access
    sudo mkdir -p /etc/systemd/system/ollama.service.d/
    sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null <<EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
EOF
    
    sudo systemctl daemon-reload
    sudo systemctl enable ollama
    sudo systemctl start ollama
    
    # Wait and download models
    sleep 10
    ollama pull llama2:latest
    ollama pull mistral:latest
fi

# Create web server directory
WEB_DIR="/home/ec2-user/excel-web-addin"
mkdir -p $WEB_DIR
cd $WEB_DIR

# Create package.json
cat > package.json << EOF
{
  "name": "excel-ollama-web-addin",
  "version": "1.0.0",
  "description": "Web-based Excel add-in for Ollama AI analysis",
  "main": "server.js",
  "scripts": {
    "start": "node server.js",
    "dev": "nodemon server.js"
  },
  "dependencies": {
    "express": "^4.18.2",
    "cors": "^2.8.5",
    "axios": "^1.6.0",
    "helmet": "^7.1.0"
  },
  "devDependencies": {
    "nodemon": "^3.0.1"
  }
}
EOF

# Install dependencies
echo "Installing Node.js dependencies..."
npm install

# Create web server
cat > server.js << EOF
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const helmet = require('helmet');
const path = require('path');

const app = express();
const PORT = 3000;
const OLLAMA_URL = 'http://localhost:11434';

// Middleware
app.use(helmet({
    contentSecurityPolicy: false, // Disable for Office.js compatibility
    crossOriginEmbedderPolicy: false
}));

app.use(cors({
    origin: ['https://localhost:3000', 'https://excel.officeapps.live.com', 'https://excel.office.com'],
    credentials: true
}));

app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// Health check
app.get('/health', (req, res) => {
    res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Ollama proxy endpoints
app.get('/api/ollama/models', async (req, res) => {
    try {
        const response = await axios.get(\`\${OLLAMA_URL}/api/tags\`);
        res.json(response.data);
    } catch (error) {
        res.status(500).json({ error: 'Failed to fetch models', details: error.message });
    }
});

app.post('/api/ollama/analyze', async (req, res) => {
    try {
        const { data, analysisType, model = 'llama2:latest' } = req.body;
        
        const prompt = createAnalysisPrompt(data, analysisType);
        
        const response = await axios.post(\`\${OLLAMA_URL}/api/generate\`, {
            model: model,
            prompt: prompt,
            stream: false
        }, {
            timeout: 120000 // 2 minutes
        });
        
        res.json({
            success: true,
            analysis: response.data.response,
            model: model,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        res.status(500).json({ 
            success: false, 
            error: 'Analysis failed', 
            details: error.message 
        });
    }
});

app.post('/api/ollama/query', async (req, res) => {
    try {
        const { data, question, model = 'llama2:latest' } = req.body;
        
        const prompt = \`
        Based on this data with \${data.rows} rows and \${data.columns} columns:
        
        Column headers: \${data.headers.join(', ')}
        
        Sample data (first 3 rows):
        \${JSON.stringify(data.sample, null, 2)}
        
        Question: \${question}
        
        Please provide a clear, actionable answer based on the data characteristics.
        Focus on practical insights and recommendations.
        \`;
        
        const response = await axios.post(\`\${OLLAMA_URL}/api/generate\`, {
            model: model,
            prompt: prompt,
            stream: false
        }, {
            timeout: 120000
        });
        
        res.json({
            success: true,
            answer: response.data.response,
            question: question,
            model: model,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        res.status(500).json({ 
            success: false, 
            error: 'Query failed', 
            details: error.message 
        });
    }
});

// Helper function to create analysis prompts
function createAnalysisPrompt(data, analysisType) {
    const baseInfo = \`
    Analyze this dataset with \${data.rows} rows and \${data.columns} columns.
    
    Column headers: \${data.headers.join(', ')}
    
    Sample data (first 5 rows):
    \${JSON.stringify(data.sample, null, 2)}
    \`;
    
    switch (analysisType) {
        case 'statistical':
            return baseInfo + \`
            Provide a comprehensive statistical analysis including:
            1. Summary statistics for numeric columns
            2. Data quality assessment (missing values, outliers)
            3. Key patterns and correlations
            4. Business insights and recommendations
            Format the response clearly with sections and bullet points.
            \`;
            
        case 'trends':
            return baseInfo + \`
            Analyze trends and patterns in this data:
            1. Identify time-based trends if applicable
            2. Detect increasing/decreasing patterns
            3. Seasonal or cyclical patterns
            4. Forecast insights and predictions
            Provide actionable business recommendations.
            \`;
            
        case 'patterns':
            return baseInfo + \`
            Detect patterns and anomalies:
            1. Unusual values or outliers
            2. Correlations between variables
            3. Clustering or grouping patterns
            4. Data quality issues
            Highlight the most important findings.
            \`;
            
        default:
            return baseInfo + 'Provide a comprehensive analysis of this data with key insights and recommendations.';
    }
}

// Start server
app.listen(PORT, '0.0.0.0', () => {
    console.log(\`Excel-Ollama Web Add-in Server running on port \${PORT}\`);
    console.log(\`Access URL: http://$EC2_PUBLIC_IP:\${PORT}\`);
    console.log(\`Ollama URL: \${OLLAMA_URL}\`);
});
EOF

# Create public directory for web files
mkdir -p public

# Create the main Excel add-in HTML file
cat > public/index.html << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Excel-Ollama AI Plugin</title>
    
    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #0078d4;
        }
        
        .header h1 {
            color: #0078d4;
            margin: 0;
        }
        
        .button-group {
            display: flex;
            gap: 10px;
            margin: 20px 0;
            flex-wrap: wrap;
        }
        
        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
            flex: 1;
            min-width: 150px;
        }
        
        .btn-primary {
            background-color: #0078d4;
            color: white;
        }
        
        .btn-primary:hover {
            background-color: #106ebe;
        }
        
        .btn-secondary {
            background-color: #6c757d;
            color: white;
        }
        
        .btn-secondary:hover {
            background-color: #5a6268;
        }
        
        .btn-success {
            background-color: #28a745;
            color: white;
        }
        
        .btn-success:hover {
            background-color: #218838;
        }
        
        .status {
            padding: 15px;
            margin: 15px 0;
            border-radius: 4px;
            font-weight: bold;
        }
        
        .status.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .status.info {
            background-color: #d1ecf1;
            color: #0c5460;
            border: 1px solid #bee5eb;
        }
        
        .loading {
            text-align: center;
            padding: 20px;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #0078d4;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .hidden {
            display: none;
        }
        
        .config-section {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 4px;
            margin: 20px 0;
        }
        
        .form-group {
            margin: 15px 0;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
        }
        
        .form-group input, .form-group select {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
        }
        
        .results {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 4px;
            margin: 20px 0;
            white-space: pre-wrap;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            max-height: 400px;
            overflow-y: auto;
        }
    </style>
</head>

<body>
    <div class="container">
        <div class="header">
            <h1>🤖 Ollama AI Analysis</h1>
            <p>AI-powered data analysis for Excel - No Python Required!</p>
        </div>
        
        <div id="status" class="status info">
            Ready to analyze your data
        </div>
        
        <div id="loading" class="loading hidden">
            <div class="spinner"></div>
            <p>Processing your request...</p>
        </div>
        
        <div class="button-group">
            <button class="btn btn-primary" onclick="analyzeData()">
                📊 Analyze Data
            </button>
            <button class="btn btn-primary" onclick="askQuestion()">
                ❓ Ask Question
            </button>
            <button class="btn btn-success" onclick="generateReport()">
                📋 Generate Report
            </button>
            <button class="btn btn-secondary" onclick="showConfig()">
                ⚙️ Configure
            </button>
        </div>
        
        <div id="config" class="config-section hidden">
            <h3>Configuration</h3>
            <div class="form-group">
                <label for="serverUrl">Ollama Server URL:</label>
                <input type="text" id="serverUrl" placeholder="http://your-ec2-ip:11434">
            </div>
            <div class="form-group">
                <label for="defaultModel">Default Model:</label>
                <select id="defaultModel">
                    <option value="llama2:latest">llama2:latest</option>
                    <option value="mistral:latest">mistral:latest</option>
                    <option value="codellama:latest">codellama:latest</option>
                    <option value="phi:latest">phi:latest</option>
                </select>
            </div>
            <div class="button-group">
                <button class="btn btn-primary" onclick="testConnection()">Test Connection</button>
                <button class="btn btn-success" onclick="saveConfig()">Save Config</button>
            </div>
        </div>
        
        <div id="results" class="results hidden"></div>
    </div>
    
    <script src="excel-addin.js"></script>
</body>
</html>
EOF

# Create the Excel add-in JavaScript file
cat > public/excel-addin.js << 'EOF'
// Excel-Ollama AI Plugin - Pure JavaScript Implementation
// No Python required on Windows!

let config = {
    serverUrl: window.location.origin, // Use current server
    defaultModel: 'llama2:latest'
};

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Excel-Ollama AI Plugin loaded successfully!');
        updateStatus('Excel-Ollama AI Plugin loaded successfully!', 'success');
        loadConfig();
    }
});

// Configuration functions
function loadConfig() {
    const savedConfig = localStorage.getItem('ollamaConfig');
    if (savedConfig) {
        config = { ...config, ...JSON.parse(savedConfig) };
        document.getElementById('serverUrl').value = config.serverUrl;
        document.getElementById('defaultModel').value = config.defaultModel;
    }
}

function saveConfig() {
    config.serverUrl = document.getElementById('serverUrl').value || config.serverUrl;
    config.defaultModel = document.getElementById('defaultModel').value;
    
    localStorage.setItem('ollamaConfig', JSON.stringify(config));
    updateStatus('Configuration saved successfully!', 'success');
    hideConfig();
}

function showConfig() {
    document.getElementById('config').classList.remove('hidden');
}

function hideConfig() {
    document.getElementById('config').classList.add('hidden');
}

// Main analysis functions
async function analyzeData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            
            await context.sync();
            
            if (range.rowCount < 2) {
                updateStatus('Please select a data range with at least 2 rows (including headers)', 'error');
                return;
            }
            
            showLoading('Analyzing your data...');
            
            const data = processExcelData(range.values);
            const analysis = await performAnalysis(data, 'statistical');
            
            await writeResultsToSheet(context, analysis, 'AI_Analysis_Results');
            
            hideLoading();
            updateStatus('Analysis completed! Check the AI_Analysis_Results sheet.', 'success');
            showResults(analysis);
        });
    } catch (error) {
        hideLoading();
        updateStatus('Analysis failed: ' + error.message, 'error');
    }
}

async function askQuestion() {
    try {
        const question = prompt('Ask a question about your data:');
        if (!question) return;
        
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (range.rowCount < 2) {
                updateStatus('Please select a data range first', 'error');
                return;
            }
            
            showLoading('Processing your question...');
            
            const data = processExcelData(range.values);
            const response = await askOllamaQuestion(data, question);
            
            await writeResultsToSheet(context, response, 'AI_Query_Results');
            
            hideLoading();
            updateStatus('Query completed! Check the AI_Query_Results sheet.', 'success');
            showResults(response);
        });
    } catch (error) {
        hideLoading();
        updateStatus('Query failed: ' + error.message, 'error');
    }
}

async function generateReport() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(['values', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (range.rowCount < 2) {
                updateStatus('Please select a data range for report generation', 'error');
                return;
            }
            
            showLoading('Generating comprehensive report...');
            
            const data = processExcelData(range.values);
            
            // Perform multiple analyses
            const analyses = await Promise.all([
                performAnalysis(data, 'statistical'),
                performAnalysis(data, 'trends'),
                performAnalysis(data, 'patterns')
            ]);
            
            const report = {
                title: 'Comprehensive Data Analysis Report',
                generated_at: new Date().toISOString(),
                analyses: analyses,
                summary: 'Multi-faceted analysis completed successfully'
            };
            
            await writeResultsToSheet(context, report, 'AI_Comprehensive_Report');
            
            hideLoading();
            updateStatus('Report generated! Check the AI_Comprehensive_Report sheet.', 'success');
            showResults(report);
        });
    } catch (error) {
        hideLoading();
        updateStatus('Report generation failed: ' + error.message, 'error');
    }
}

async function testConnection() {
    try {
        showLoading('Testing connection...');
        
        const response = await fetch(`${config.serverUrl}/api/ollama/models`);
        const data = await response.json();
        
        if (response.ok && data.models) {
            hideLoading();
            updateStatus(`Connection successful! Found ${data.models.length} models.`, 'success');
        } else {
            throw new Error('Invalid response from server');
        }
    } catch (error) {
        hideLoading();
        updateStatus('Connection failed: ' + error.message, 'error');
    }
}

// Data processing functions
function processExcelData(values) {
    if (!values || values.length < 2) return null;
    
    const headers = values[0];
    const rows = values.slice(1);
    
    const processedRows = rows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });
    
    return {
        headers: headers,
        rows: rows.length,
        columns: headers.length,
        sample: processedRows.slice(0, 5), // First 5 rows as sample
        data: processedRows
    };
}

// API functions
async function performAnalysis(data, analysisType) {
    try {
        const response = await fetch(`${config.serverUrl}/api/ollama/analyze`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                data: data,
                analysisType: analysisType,
                model: config.defaultModel
            })
        });
        
        const result = await response.json();
        
        if (!response.ok) {
            throw new Error(result.error || 'Analysis failed');
        }
        
        return {
            type: analysisType,
            result: result.analysis,
            model: result.model,
            timestamp: result.timestamp
        };
    } catch (error) {
        throw new Error(`Analysis failed: ${error.message}`);
    }
}

async function askOllamaQuestion(data, question) {
    try {
        const response = await fetch(`${config.serverUrl}/api/ollama/query`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                data: data,
                question: question,
                model: config.defaultModel
            })
        });
        
        const result = await response.json();
        
        if (!response.ok) {
            throw new Error(result.error || 'Query failed');
        }
        
        return {
            question: question,
            answer: result.answer,
            model: result.model,
            timestamp: result.timestamp
        };
    } catch (error) {
        throw new Error(`Query failed: ${error.message}`);
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
    
    // Format results for Excel
    const resultText = formatResultsForExcel(results);
    const lines = resultText.split('\n');
    const values = lines.map(line => [line]);
    
    if (values.length > 0) {
        const range = sheet.getRange(`A1:A${lines.length}`);
        range.values = values;
        
        // Format the sheet
        range.format.font.name = 'Consolas';
        range.format.font.size = 10;
        
        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns();
    }
    
    // Activate the sheet
    sheet.activate();
    
    await context.sync();
}

function formatResultsForExcel(results) {
    if (typeof results === 'string') {
        return results;
    }
    
    let formatted = '';
    
    if (results.title) {
        formatted += `${results.title}\n`;
        formatted += '='.repeat(results.title.length) + '\n\n';
    }
    
    if (results.question) {
        formatted += `Question: ${results.question}\n\n`;
        formatted += `Answer: ${results.answer}\n\n`;
    }
    
    if (results.result) {
        formatted += `Analysis Result:\n${results.result}\n\n`;
    }
    
    if (results.analyses) {
        results.analyses.forEach((analysis, index) => {
            formatted += `Analysis ${index + 1} (${analysis.type}):\n`;
            formatted += `${analysis.result}\n\n`;
        });
    }
    
    if (results.model) {
        formatted += `Model Used: ${results.model}\n`;
    }
    
    if (results.timestamp) {
        formatted += `Generated: ${new Date(results.timestamp).toLocaleString()}\n`;
    }
    
    return formatted || JSON.stringify(results, null, 2);
}

// UI helper functions
function updateStatus(message, type = 'info') {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
}

function showLoading(message = 'Processing...') {
    const loadingDiv = document.getElementById('loading');
    loadingDiv.querySelector('p').textContent = message;
    loadingDiv.classList.remove('hidden');
}

function hideLoading() {
    document.getElementById('loading').classList.add('hidden');
}

function showResults(results) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.textContent = formatResultsForExcel(results);
    resultsDiv.classList.remove('hidden');
}
EOF

# Create manifest.xml for Excel add-in
cat > public/manifest.xml << EOF
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
           xsi:type="TaskPaneApp">

  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Excel-Ollama AI Plugin</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  
  <DisplayName DefaultValue="Ollama AI Analysis" />
  <Description DefaultValue="AI-powered data analysis using Ollama models. No Python required!" />
  
  <IconUrl DefaultValue="http://$EC2_PUBLIC_IP:3000/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="http://$EC2_PUBLIC_IP:3000/icon-64.png" />
  
  <SupportUrl DefaultValue="https://github.com/your-repo/excel-ollama-plugin" />
  
  <AppDomains>
    <AppDomain>http://$EC2_PUBLIC_IP:3000</AppDomain>
  </AppDomains>
  
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  
  <DefaultSettings>
    <SourceLocation DefaultValue="http://$EC2_PUBLIC_IP:3000/index.html" />
  </DefaultSettings>
  
  <Permissions>ReadWriteDocument</Permissions>

</OfficeApp>
EOF

# Create simple icons
mkdir -p public/icons
echo "Creating placeholder icons..."

# Create a simple SVG icon and convert to PNG (placeholder)
cat > public/icon-32.png << 'EOF'
# This would be a 32x32 PNG icon
# For now, we'll create a placeholder
EOF

cat > public/icon-64.png << 'EOF'
# This would be a 64x64 PNG icon  
# For now, we'll create a placeholder
EOF

# Create systemd service for auto-start
sudo tee /etc/systemd/system/excel-web-addin.service > /dev/null << EOF
[Unit]
Description=Excel-Ollama Web Add-in Server
After=network.target

[Service]
Type=simple
User=ec2-user
WorkingDirectory=$WEB_DIR
ExecStart=/usr/bin/node server.js
Restart=always
RestartSec=10
Environment=NODE_ENV=production

[Install]
WantedBy=multi-user.target
EOF

# Enable and start the service
sudo systemctl daemon-reload
sudo systemctl enable excel-web-addin
sudo systemctl start excel-web-addin

# Configure firewall for port 3000
if command -v ufw &> /dev/null; then
    sudo ufw allow 3000/tcp
fi

echo ""
echo "=========================================="
echo "Web-based Excel Add-in Setup Complete!"
echo "=========================================="
echo "Server URL: http://$EC2_PUBLIC_IP:3000"
echo "Manifest URL: http://$EC2_PUBLIC_IP:3000/manifest.xml"
echo "Ollama URL: http://$EC2_PUBLIC_IP:11434"
echo ""
echo "Available models:"
ollama list
echo ""
echo "Service status:"
sudo systemctl status excel-web-addin --no-pager
echo ""
echo "NEXT STEPS FOR WINDOWS:"
echo "1. Open Excel on your Windows laptop"
echo "2. Go to Insert > Get Add-ins > Upload My Add-in"
echo "3. Upload the manifest.xml file from: http://$EC2_PUBLIC_IP:3000/manifest.xml"
echo "4. The add-in will appear in Excel's task pane"
echo "5. Configure the server URL and start analyzing!"
echo ""
echo "IMPORTANT: Update your EC2 Security Group to allow:"
echo "- Port 3000 (Web server)"
echo "- Port 11434 (Ollama)"
echo "- Source: Your laptop's public IP"
echo "=========================================="
EOF

chmod +x web-server-setup.sh