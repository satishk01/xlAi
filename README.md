# Excel-Ollama AI Plugin

A comprehensive Excel plugin that integrates with Ollama models to perform intelligent, agentic activity data analysis. Transform Excel into a powerful AI-enhanced data analysis platform using local LLM capabilities.

## Features

### 🤖 AI-Powered Analysis
- **Statistical Analysis**: Comprehensive descriptive statistics and correlations
- **Trend Analysis**: Time series analysis with forecasting capabilities
- **Pattern Detection**: Seasonal patterns, anomalies, and outlier detection
- **Clustering**: Automatic grouping of similar data points
- **Natural Language Queries**: Ask questions about your data in plain English

### 📊 Excel Integration
- **Custom Ribbon**: Dedicated "Ollama AI Analysis" tab with intuitive controls
- **Excel Functions**: Custom UDFs like `=OLLAMA_ANALYZE()`, `=AI_TREND()`, `=PATTERN_DETECT()`
- **Automatic Reporting**: Generate comprehensive reports and dashboards
- **Visualization**: Create charts and visualizations based on analysis results
- **Export Capabilities**: Export results in multiple formats

### 🔧 Ollama Integration
- **Multiple Models**: Support for llama2, codellama, mistral, phi, and more
- **Local Processing**: All data stays on your machine - complete privacy
- **Streaming Responses**: Real-time analysis feedback
- **Model Switching**: Change models without restarting Excel
- **Custom Parameters**: Configure model settings for optimal performance

## Prerequisites

- **Microsoft Excel 2016 or later**
- **Python 3.8 or later**
- **Ollama installed and running**
- **Windows 10 or later**

## Installation

### Option A: Local Installation

#### 1. Install Ollama Locally

```bash
# Download and install Ollama from https://ollama.ai
# Then start the server
ollama serve

# Download some models
ollama pull llama2
ollama pull codellama
ollama pull mistral
```

### Option B: Cloud Deployment (Recommended)

#### 1. Deploy Ollama on Amazon EC2

For better performance and to avoid using local resources:

```bash
# 1. Launch EC2 instance (t3.medium or larger)
# 2. Configure Security Group (port 11434)
# 3. Run setup script on EC2:
curl -O https://raw.githubusercontent.com/your-repo/excel-ollama-plugin/main/ec2-setup.sh
chmod +x ec2-setup.sh
sudo ./ec2-setup.sh
```

**📖 See [QUICK_START_GUIDE.md](QUICK_START_GUIDE.md) for detailed EC2 setup instructions.**

### 2. Install the Plugin

#### Option A: Automatic Installation (Recommended)

1. Download the plugin package
2. Run the installer:

```bash
python install.py --install
```

#### Option B: Manual Installation

1. Clone or download this repository
2. Install Python dependencies:

```bash
pip install -r requirements.txt
```

3. Run the setup script:

```bash
python setup.py install
```

4. Register the Excel add-in:

```bash
python -c "import xlwings as xw; xw.addin install"
```

### 3. Configure the Plugin

1. Open Excel
2. Look for the "Ollama AI Analysis" tab in the ribbon
3. Click "Configure" to set up your Ollama connection:
   - **Local:** `http://localhost:11434`
   - **EC2:** `http://YOUR_EC2_PUBLIC_IP:11434`
4. Test the connection and select your preferred model

**🧪 Test your connection:** `python test-connection.py http://YOUR_SERVER:11434`

## Quick Start

### Basic Analysis

1. **Select your data** in Excel (including headers)
2. **Click "Analyze Data"** in the Ollama AI Analysis ribbon
3. **Review results** in the automatically generated sheet
4. **Explore insights** in the results dialog

### Natural Language Queries

1. **Select your data range**
2. **Click "Ask Question"**
3. **Type your question** (e.g., "What trends do you see in sales data?")
4. **Get AI-powered insights** in natural language

### Using Custom Functions

Add AI analysis directly to your spreadsheet:

```excel
=OLLAMA_ANALYZE(A1:D100, "Analyze sales trends")
=AI_TREND(B1:B50, 10)
=PATTERN_DETECT(C1:C200, 0.8)
```

## Usage Examples

### Example 1: Sales Data Analysis

```excel
# Select your sales data (Date, Product, Sales, Region)
# Click "Trend Analysis"
# Get insights like:
# - "Sales show strong upward trend (R² = 0.85)"
# - "Peak sales occur on Fridays and Saturdays"
# - "Q4 shows 23% higher sales than average"
```

### Example 2: Customer Behavior Analysis

```excel
# Select customer activity data
# Ask: "What patterns do you see in customer behavior?"
# Get responses like:
# - "Customers are most active between 2-4 PM"
# - "Weekend activity is 40% higher than weekdays"
# - "Detected 3 distinct customer segments"
```

### Example 3: Financial Forecasting

```excel
# Select time series financial data
# Use: =AI_TREND(A1:B100, 12)
# Get 12-period forecast with confidence intervals
```

## Configuration Options

### Ollama Settings
- **Server URL**: Default `http://localhost:11434`
- **Default Model**: Choose from available models
- **Timeout**: Request timeout in seconds
- **Streaming**: Enable real-time responses

### Analysis Settings
- **Default Analysis Type**: Statistical, trend, or pattern analysis
- **Max Rows**: Maximum rows to process (default: 100,000)
- **Chunk Size**: Size for processing large datasets
- **Auto-detect Types**: Automatically detect data types

### UI Preferences
- **Show Progress**: Display progress indicators
- **Auto-open Results**: Automatically open results sheets
- **Show Confidence**: Display confidence scores
- **Notifications**: Enable completion and error notifications

## Troubleshooting

### Common Issues

#### "Cannot connect to Ollama"
- Ensure Ollama is running: `ollama serve`
- Check server URL in configuration
- Verify firewall settings

#### "Model not found"
- Download the model: `ollama pull llama2`
- Refresh models in configuration
- Check model name spelling

#### "Analysis takes too long"
- Reduce data size or use sampling
- Enable chunking for large datasets
- Check system resources

#### "Excel functions not working"
- Restart Excel after installation
- Check if plugin is enabled
- Verify xlwings installation

### Getting Help

1. **Check the logs**: Located in `%LOCALAPPDATA%\ExcelOllamaPlugin\logs\`
2. **Use "Test Connection"** in the configuration dialog
3. **Review the help documentation** in the plugin
4. **Check system requirements** and prerequisites

## Advanced Usage

### Custom Analysis Workflows

Create complex analysis pipelines:

1. **Statistical Overview** → Basic data understanding
2. **Pattern Detection** → Identify anomalies and trends  
3. **Clustering** → Group similar data points
4. **Forecasting** → Predict future values
5. **Report Generation** → Comprehensive insights

### Batch Processing

Process multiple datasets:

```python
# Use the plugin's Python API for batch processing
from src.main import ExcelOllamaPlugin

plugin = ExcelOllamaPlugin()
results = []

for sheet in workbook.sheets:
    data = plugin.excel_interface.get_worksheet_data(sheet.name)
    result = await plugin.agent_controller.execute_analysis_pipeline(
        data, 'statistical_analysis'
    )
    results.append(result)
```

### Custom Models

Use specialized Ollama models:

1. Download custom models: `ollama pull your-custom-model`
2. Select in the model dropdown
3. Configure parameters for optimal performance

## API Reference

### Custom Excel Functions

#### `OLLAMA_ANALYZE(range, prompt)`
- **range**: Data range to analyze
- **prompt**: Analysis instruction
- **Returns**: Analysis summary

#### `AI_TREND(range, periods)`
- **range**: Time series data
- **periods**: Number of periods to forecast
- **Returns**: Trend analysis and forecast

#### `PATTERN_DETECT(range, threshold)`
- **range**: Data to analyze for patterns
- **threshold**: Detection sensitivity (0-1)
- **Returns**: Pattern detection results

### Python API

```python
from src.core.agent_controller import AgentController
from src.core.ollama_client import OllamaClient

# Initialize components
ollama_client = OllamaClient("http://localhost:11434")
agent_controller = AgentController(ollama_client)

# Run analysis
result = await agent_controller.execute_analysis_pipeline(
    data, 'trend_analysis', {'periods': 10}
)
```

## Development

### Project Structure

```
excel-ollama-ai-plugin/
├── src/
│   ├── core/                 # Core functionality
│   │   ├── ollama_client.py  # Ollama API integration
│   │   ├── agent_controller.py # Agent coordination
│   │   ├── excel_interface.py # Excel integration
│   │   └── query_processor.py # Natural language processing
│   ├── agents/               # AI agents
│   │   ├── analysis_agent.py # Statistical analysis
│   │   ├── pattern_agent.py  # Pattern detection
│   │   └── reporting_agent.py # Report generation
│   ├── ui/                   # User interface
│   │   ├── ribbon_ui.xml     # Excel ribbon definition
│   │   └── dialog_forms.py   # Configuration dialogs
│   └── utils/                # Utilities
├── tests/                    # Test suite
├── manifest.xml              # Excel add-in manifest
├── requirements.txt          # Python dependencies
├── setup.py                  # Package setup
└── install.py               # Installation script
```

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

### Testing

Run the test suite:

```bash
python -m pytest tests/
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- **Documentation**: Check the built-in help system
- **Issues**: Report bugs and feature requests on GitHub
- **Community**: Join our discussions and share your use cases

## Changelog

### Version 1.0.0
- Initial release
- Core analysis capabilities
- Excel integration
- Ollama model support
- Natural language queries
- Comprehensive reporting

---

**Transform your Excel data analysis with the power of local AI models!** 🚀