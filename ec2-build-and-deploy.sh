#!/bin/bash
# Complete EC2 Build and Deploy Script for Excel-Ollama AI Plugin
# This script builds the plugin on EC2 and creates a Windows deployment package

echo "Excel-Ollama AI Plugin - EC2 Build & Deploy"
echo "============================================"

# Set variables
PLUGIN_NAME="ExcelOllamaAIPlugin"
VERSION="1.0.0"
BUILD_DIR="/home/ec2-user/excel-ollama-build"
DEPLOY_DIR="/home/ec2-user/excel-ollama-deploy"

# Get EC2 public IP
EC2_PUBLIC_IP=$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)
echo "EC2 Public IP: $EC2_PUBLIC_IP"

# Update system
echo "Updating system packages..."
sudo yum update -y || sudo apt update -y

# Install required packages
echo "Installing build dependencies..."
if command -v yum &> /dev/null; then
    # Amazon Linux / RHEL / CentOS
    sudo yum install -y curl wget git python3 python3-pip zip unzip nodejs npm
elif command -v apt &> /dev/null; then
    # Ubuntu / Debian
    sudo apt install -y curl wget git python3 python3-pip zip unzip nodejs npm
fi

# Install Python build tools
pip3 install --user wheel setuptools pyinstaller

# Create build directory
echo "Setting up build environment..."
mkdir -p $BUILD_DIR
mkdir -p $DEPLOY_DIR
cd $BUILD_DIR

# Clone or copy plugin source code
echo "Getting plugin source code..."
# If you have the code in a git repo:
# git clone https://github.com/your-repo/excel-ollama-plugin.git .

# For now, we'll assume the code is already present
# Copy the source code to build directory if needed

# Install Python dependencies
echo "Installing Python dependencies..."
pip3 install --user -r requirements.txt

# Install Ollama
echo "Installing Ollama..."
curl -fsSL https://ollama.ai/install.sh | sh

# Configure Ollama for external access
echo "Configuring Ollama..."
sudo mkdir -p /etc/systemd/system/ollama.service.d/
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null <<EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
EOF

# Start Ollama service
sudo systemctl daemon-reload
sudo systemctl enable ollama
sudo systemctl start ollama

# Wait for Ollama to start
echo "Waiting for Ollama to initialize..."
sleep 10

# Download AI models
echo "Downloading AI models (this may take 15-20 minutes)..."
ollama pull llama2:latest
ollama pull codellama:latest
ollama pull mistral:latest
ollama pull phi:latest

# Build the plugin package
echo "Building plugin package..."
python3 setup.py bdist_wheel

# Create Windows deployment package
echo "Creating Windows deployment package..."
cd $DEPLOY_DIR

# Create the deployment structure
mkdir -p windows-plugin/{plugin,scripts,config,docs}

# Copy built plugin
cp -r $BUILD_DIR/src windows-plugin/plugin/
cp $BUILD_DIR/requirements.txt windows-plugin/
cp $BUILD_DIR/manifest.xml windows-plugin/
cp $BUILD_DIR/setup.py windows-plugin/

# Create Windows-specific installer
cat > windows-plugin/install-plugin.bat << 'EOF'
@echo off
echo Excel-Ollama AI Plugin - Windows Installation
echo =============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python 3.8+ from: https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found. Installing plugin...
echo.

REM Install Python dependencies
echo Installing Python dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

REM Install xlwings Excel add-in
echo Installing Excel add-in...
python -m pip install xlwings
python -c "import xlwings as xw; xw.addin install"

REM Copy plugin files to user directory
echo Installing plugin files...
set PLUGIN_DIR=%APPDATA%\ExcelOllamaPlugin
mkdir "%PLUGIN_DIR%" 2>nul
xcopy /E /Y plugin "%PLUGIN_DIR%\plugin\"
copy config.json "%PLUGIN_DIR%\" 2>nul

REM Create Excel startup file
echo Creating Excel startup integration...
set EXCEL_STARTUP=%APPDATA%\Microsoft\Excel\XLSTART
mkdir "%EXCEL_STARTUP%" 2>nul

(
echo import sys
echo import os
echo sys.path.insert(0, os.path.join(os.getenv('APPDATA'^), 'ExcelOllamaPlugin'^)^)
echo try:
echo     from plugin.main import initialize_plugin
echo     plugin = initialize_plugin(^)
echo     print("Excel-Ollama AI Plugin loaded successfully"^)
echo except Exception as e:
echo     print(f"Failed to load plugin: {e}"^)
) > "%EXCEL_STARTUP%\ExcelOllamaPlugin.py"

echo.
echo =============================================
echo Installation completed successfully!
echo =============================================
echo.
echo IMPORTANT: Your Ollama server is running at:
echo http://EC2_PUBLIC_IP_PLACEHOLDER:11434
echo.
echo Next steps:
echo 1. Restart Excel
echo 2. Look for "Ollama AI Analysis" tab in the ribbon
echo 3. Click "Configure" and set server URL to the above address
echo 4. Test the connection
echo 5. Start analyzing your data!
echo.
pause
EOF

# Replace placeholder with actual IP
sed -i "s/EC2_PUBLIC_IP_PLACEHOLDER/$EC2_PUBLIC_IP/g" windows-plugin/install-plugin.bat

# Create configuration file with EC2 IP
cat > windows-plugin/config.json << EOF
{
  "ollama": {
    "server_url": "http://$EC2_PUBLIC_IP:11434",
    "default_model": "llama2:latest",
    "timeout": 300,
    "max_retries": 3,
    "stream_responses": true,
    "connection_test_timeout": 10
  },
  "excel_settings": {
    "auto_refresh": true,
    "max_rows_per_chunk": 10000,
    "default_chart_type": "line",
    "enable_custom_functions": true
  },
  "ui_preferences": {
    "show_progress_dialogs": true,
    "auto_save_results": true,
    "show_tooltips": true
  }
}
EOF

# Create uninstaller
cat > windows-plugin/uninstall-plugin.bat << 'EOF'
@echo off
echo Excel-Ollama AI Plugin - Uninstaller
echo ====================================
echo.

echo Removing plugin files...
rmdir /S /Q "%APPDATA%\ExcelOllamaPlugin" 2>nul
del "%APPDATA%\Microsoft\Excel\XLSTART\ExcelOllamaPlugin.py" 2>nul

echo Uninstalling xlwings add-in...
python -c "import xlwings as xw; xw.addin remove" 2>nul

echo.
echo Plugin uninstalled successfully!
echo Please restart Excel to complete the removal.
echo.
pause
EOF

# Create connection test script
cat > windows-plugin/test-connection.bat << EOF
@echo off
echo Testing connection to Ollama server...
echo.
curl -s http://$EC2_PUBLIC_IP:11434/api/tags
if errorlevel 1 (
    echo ERROR: Cannot connect to Ollama server
    echo Make sure the EC2 instance is running
) else (
    echo SUCCESS: Connection to Ollama server working!
)
echo.
pause
EOF

# Create README for Windows users
cat > windows-plugin/README-WINDOWS.txt << EOF
Excel-Ollama AI Plugin - Windows Installation Package
====================================================

This package contains everything needed to use the Excel-Ollama AI Plugin
with your Ollama server running on EC2.

CONTENTS:
- install-plugin.bat     : Main installer script
- uninstall-plugin.bat   : Uninstaller script  
- test-connection.bat    : Test connection to EC2 server
- plugin/                : Plugin source code
- config.json           : Pre-configured settings
- requirements.txt      : Python dependencies

INSTALLATION:
1. Make sure Python 3.8+ is installed on Windows
2. Double-click "install-plugin.bat"
3. Follow the prompts
4. Restart Excel
5. Look for "Ollama AI Analysis" tab

OLLAMA SERVER:
Your Ollama server is running at: http://$EC2_PUBLIC_IP:11434

USAGE:
1. Open Excel with your data
2. Select data range
3. Go to "Ollama AI Analysis" tab
4. Click "Analyze Data" or "Ask Question"
5. Get AI-powered insights!

TROUBLESHOOTING:
- Run "test-connection.bat" to verify server connectivity
- Check that EC2 instance is running
- Verify Security Group allows port 11434 from your IP
- Restart Excel if ribbon tab doesn't appear

SUPPORT:
- Check the logs in: %APPDATA%\ExcelOllamaPlugin\logs\
- Ensure your IP is allowed in EC2 Security Group
- Verify Ollama service is running on EC2

Build Date: $(date)
EC2 Server: http://$EC2_PUBLIC_IP:11434
EOF

# Create ZIP package for easy download
echo "Creating deployment package..."
zip -r "ExcelOllamaPlugin-Windows-v$VERSION.zip" windows-plugin/

# Create download instructions
cat > download-instructions.txt << EOF
Excel-Ollama AI Plugin - Download Instructions
==============================================

Your plugin has been built and is ready for download!

DOWNLOAD THE PLUGIN:
1. From your Windows laptop, download the ZIP file:
   scp -i your-key.pem ec2-user@$EC2_PUBLIC_IP:$DEPLOY_DIR/ExcelOllamaPlugin-Windows-v$VERSION.zip .

2. Or use WinSCP/FileZilla to download:
   File: $DEPLOY_DIR/ExcelOllamaPlugin-Windows-v$VERSION.zip

INSTALLATION ON WINDOWS:
1. Extract the ZIP file
2. Double-click "install-plugin.bat"
3. Follow the prompts
4. Restart Excel

SERVER INFORMATION:
- Ollama Server URL: http://$EC2_PUBLIC_IP:11434
- Available Models: $(ollama list | grep -v NAME | awk '{print $1}' | tr '\n' ', ' | sed 's/,$//')

SECURITY:
Make sure your Windows laptop IP is allowed in the EC2 Security Group:
- Port: 11434
- Protocol: TCP
- Source: Your laptop's public IP

The plugin is ready to use!
EOF

echo ""
echo "============================================"
echo "BUILD COMPLETED SUCCESSFULLY!"
echo "============================================"
echo "Package location: $DEPLOY_DIR/ExcelOllamaPlugin-Windows-v$VERSION.zip"
echo "EC2 Ollama Server: http://$EC2_PUBLIC_IP:11434"
echo ""
echo "Available models:"
ollama list
echo ""
echo "To download to your Windows laptop:"
echo "scp -i your-key.pem ec2-user@$EC2_PUBLIC_IP:$DEPLOY_DIR/ExcelOllamaPlugin-Windows-v$VERSION.zip ."
echo ""
echo "Next steps:"
echo "1. Download the ZIP file to your Windows laptop"
echo "2. Extract and run install-plugin.bat"
echo "3. Restart Excel"
echo "4. Configure the plugin to use: http://$EC2_PUBLIC_IP:11434"
echo "============================================"