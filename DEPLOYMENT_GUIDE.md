# Complete Deployment Guide: EC2 + Local Excel Plugin

This guide shows you how to deploy Ollama on Amazon EC2 and use the Excel plugin on your laptop to connect to it.

## ??? Architecture Overview

```
Your Laptop (Windows)          Amazon EC2 (Linux)
???????????????????????       ???????????????????????
?  Excel + AI Plugin ? ????? ?  Ollama Server      ?
?  (Port: Any)        ?       ?  (Port: 11434)      ?
???????????????????????       ???????????????????????
```

## Part 1: Setting Up Amazon EC2 Instance

### Step 1: Launch EC2 Instance

1. **Go to AWS Console** ? EC2 ? Launch Instance
2. **Choose AMI:** Amazon Linux 2 or Ubuntu 20.04 LTS
3. **Instance Type:** 
   - **Minimum:** t3.medium (2 vCPU, 4 GB RAM)
   - **Recommended:** t3.large (2 vCPU, 8 GB RAM)
   - **For heavy workloads:** t3.xlarge or larger
4. **Storage:** At least 20 GB (models can be large)
5. **Security Group:** Create new or use existing (we'll configure later)

### Step 2: Configure Security Group

1. **Go to Security Groups** in EC2 console
2. **Add Inbound Rule:**
   ```
   Type: Custom TCP
   Port: 11434
   Source: Your laptop's public IP (find at whatismyipaddress.com)
   Description: Ollama API for Excel Plugin
   ```

### Step 3: Connect to EC2 and Install Ollama

```bash
# Connect to your EC2 instance
ssh -i your-key.pem ec2-user@your-ec2-public-ip

# Download and run the setup script
curl -O https://raw.githubusercontent.com/your-repo/excel-ollama-plugin/main/ec2-setup.sh
chmod +x ec2-setup.sh
sudo ./ec2-setup.sh
```

Or manually install:

```bash
# Update system
sudo yum update -y  # For Amazon Linux
# sudo apt update -y  # For Ubuntu

# Install Ollama
curl -fsSL https://ollama.ai/install.sh | sh

# Configure Ollama for external access
sudo mkdir -p /etc/systemd/system/ollama.service.d/
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null <<EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
EOF

# Restart Ollama
sudo systemctl daemon-reload
sudo systemctl restart ollama
sudo systemctl enable ollama

# Download models
ollama pull llama2:latest
ollama pull codellama:latest
ollama pull mistral:latest
```

### Step 4: Test EC2 Setup

```bash
# Get your EC2 public IP
curl http://169.254.169.254/latest/meta-data/public-ipv4

# Test Ollama is running
curl http://localhost:11434/api/tags

# Test external access (from your laptop)
curl http://YOUR_EC2_PUBLIC_IP:11434/api/tags

http://18.215.161.204:11434/api/tags
```

## Part 2: Building and Installing Excel Plugin

### Step 1: Prepare Your Laptop Environment

```bash
# Install Python 3.8+ (if not already installed)
# Download from https://python.org

# Install Git (if not already installed)
# Download from https://git-scm.com

# Clone or download the plugin code
git clone https://github.com/your-repo/excel-ollama-plugin.git
cd excel-ollama-plugin
```

### Step 2: Install Plugin Dependencies

```bash
# Install Python dependencies
pip install -r requirements.txt

# Key dependencies that will be installed:
# - xlwings (Excel integration)
# - pandas (data processing)
# - numpy (numerical operations)
# - requests (HTTP client)
# - aiohttp (async HTTP)
# - scikit-learn (machine learning)
# - scipy (scientific computing)
# - tkinter (GUI - usually included with Python)
```

### Step 3: Build and Install Plugin

```bash
# Option 1: Automatic installation
python install.py --install

# Option 2: Manual installation
python setup.py install

# Option 3: Development installation
pip install -e .
```

### Step 4: Register Excel Add-in

```bash
# Register with xlwings
python -c "import xlwings as xw; xw.addin install"

# Or manually copy to Excel startup folder
# Copy plugin files to: %APPDATA%\Microsoft\Excel\XLSTART\
```

## Part 3: Configuring Plugin for Remote Ollama

### Step 1: Configure Plugin for EC2

1. **Open Excel**
2. **Look for "Ollama AI Analysis" tab** in the ribbon
3. **Click "Configure"**
4. **Update settings:**
   ```
   Server URL: http://YOUR_EC2_PUBLIC_IP:11434
   Default Model: llama2:latest
   Timeout: 300 seconds
   Max Retries: 3
   ```
5. **Click "Test Connection"**
6. **Save configuration**

### Step 2: Alternative - Manual Configuration

Edit the config file directly:
```json
{
  "ollama": {
    "server_url": "http://YOUR_EC2_PUBLIC_IP:11434",
    "default_model": "llama2:latest",
    "timeout": 300,
    "max_retries": 3,
    "stream_responses": true
  }
}
```

Location: `%APPDATA%\ExcelOllamaPlugin\config.json`

## Part 4: Using the Plugin

### Step 1: Basic Usage

1. **Open Excel** with your data
2. **Select data range** (including headers)
3. **Go to "Ollama AI Analysis" tab**
4. **Click "Analyze Data"**
5. **Review results** in the generated sheet

### Step 2: Natural Language Queries

1. **Select your data**
2. **Click "Ask Question"**
3. **Type questions like:**
   - "What trends do you see in sales data?"
   - "Are there any outliers in this dataset?"
   - "Can you forecast the next 10 periods?"
4. **Get AI-powered insights**

### Step 3: Custom Excel Functions

Use these functions directly in Excel cells:

```excel
=OLLAMA_ANALYZE(A1:D100, "Analyze sales trends")
=AI_TREND(B1:B50, 10)
=PATTERN_DETECT(C1:C200, 0.8)
```

## Part 5: Troubleshooting

### Common Issues and Solutions

#### 1. "Cannot connect to Ollama"
```bash
# Check EC2 instance is running
aws ec2 describe-instances --instance-ids i-1234567890abcdef0

# Check Ollama service on EC2
ssh -i your-key.pem ec2-user@your-ec2-ip
sudo systemctl status ollama

# Test connection from laptop
curl http://YOUR_EC2_IP:11434/api/tags
```

#### 2. "Connection timeout"
- Check security group allows port 11434
- Verify your laptop's IP hasn't changed
- Increase timeout in plugin configuration

#### 3. "Model not found"
```bash
# On EC2, check available models
ollama list

# Download missing models
ollama pull llama2:latest
```

#### 4. "Excel functions not working"
- Restart Excel after installation
- Check if xlwings add-in is enabled
- Verify plugin is in Excel startup folder

### Performance Optimization

#### For Large Datasets:
1. **Enable chunking** in plugin settings
2. **Use sampling** for initial analysis
3. **Increase EC2 instance size** if needed

#### For Better Response Times:
1. **Use faster models** (phi, mistral)
2. **Enable streaming responses**
3. **Keep models loaded** (they stay in memory)

## Part 6: Security Best Practices

### 1. Network Security
```bash
# Use specific IP instead of 0.0.0.0/0
# In Security Group, set Source to: YOUR_LAPTOP_IP/32

# Optional: Use VPN for additional security
# Set up AWS Client VPN or use your corporate VPN
```

### 2. Instance Security
```bash
# Keep system updated
sudo yum update -y

# Use IAM roles instead of access keys
# Limit EC2 permissions to minimum required

# Enable CloudTrail for audit logging
```

### 3. Data Security
- Plugin processes data locally in Excel
- Only analysis requests sent to EC2
- No raw data transmitted (privacy preserved)

## Part 7: Cost Optimization

### EC2 Cost Management:
1. **Use Spot Instances** for development/testing
2. **Stop instance** when not in use (data persists)
3. **Use Reserved Instances** for production
4. **Monitor usage** with CloudWatch

### Example Monthly Costs (US East):
- **t3.medium:** ~$30/month (if running 24/7)
- **t3.large:** ~$60/month (if running 24/7)
- **Storage:** ~$2/month for 20GB

### Auto-Start/Stop Script:
```bash
# Start EC2 instance when needed
aws ec2 start-instances --instance-ids i-1234567890abcdef0

# Stop when done (saves ~70% cost)
aws ec2 stop-instances --instance-ids i-1234567890abcdef0
```

## Part 8: Advanced Configuration

### Multiple Model Support:
```bash
# On EC2, install various models
ollama pull llama2:7b
ollama pull llama2:13b
ollama pull codellama:7b
ollama pull mistral:7b
ollama pull phi:latest
```

### Load Balancing (Multiple EC2 Instances):
```python
# In plugin config, use multiple URLs
"server_urls": [
    "http://ec2-1-ip:11434",
    "http://ec2-2-ip:11434"
]
```

### HTTPS Setup (Production):
```bash
# Install nginx as reverse proxy
sudo yum install nginx -y

# Configure SSL certificate
# Use Let's Encrypt or AWS Certificate Manager

# Update plugin config to use HTTPS
"server_url": "https://your-domain.com:443"
```

## ?? You're Ready!

Your Excel-Ollama AI Plugin is now deployed and ready to use! You can:

? Analyze data with natural language queries  
? Get AI-powered insights and recommendations  
? Create automated reports and dashboards  
? Use custom Excel functions for AI analysis  
? Process large datasets efficiently  
? Switch between different AI models  

**Happy analyzing!** ??