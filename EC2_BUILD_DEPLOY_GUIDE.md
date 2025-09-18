# 🏗️ EC2 Build & Windows Deploy Guide

This guide shows how to build the Excel-Ollama AI Plugin entirely on EC2 and deploy it to your Windows laptop with Excel. **No Python development environment needed on Windows!**

## 🎯 Architecture

```
EC2 Linux Instance                    Windows Laptop
┌─────────────────────┐               ┌─────────────────────┐
│  Build Environment  │               │  Excel Only         │
│  ├── Python 3.8+    │               │  ├── Microsoft      │
│  ├── Plugin Source  │  ──────────►  │  │   Excel 2016+    │
│  ├── Build Tools    │   Deploy      │  ├── Pre-built     │
│  └── Ollama Server  │   Package     │  │   Plugin        │
└─────────────────────┘               └─────────────────────┘
```

## Part 1: EC2 Setup and Build (20 minutes)

### Step 1: Launch EC2 Instance

```bash
# In AWS Console:
# 1. Launch Instance → Amazon Linux 2
# 2. Instance Type: t3.medium (minimum) or t3.large (recommended)
# 3. Storage: 30 GB (for models and build artifacts)
# 4. Security Group: Allow SSH (22) and Ollama (11434)
# 5. Create/use Key Pair
```

### Step 2: Connect and Run Build Script

```bash
# Connect to EC2
ssh -i your-key.pem ec2-user@YOUR_EC2_PUBLIC_IP

# Download and run the complete build script
curl -O https://raw.githubusercontent.com/your-repo/excel-ollama-plugin/main/ec2-build-and-deploy.sh
chmod +x ec2-build-and-deploy.sh
sudo ./ec2-build-and-deploy.sh
```

**What the build script does:**
- ✅ Installs Python 3.8+ and build tools
- ✅ Installs and configures Ollama
- ✅ Downloads AI models (llama2, codellama, mistral, phi)
- ✅ Builds the plugin package
- ✅ Creates Windows deployment package
- ✅ Configures everything for your EC2 IP

### Step 3: Download the Built Package

After the build completes, download the package to your Windows laptop:

```bash
# From your Windows laptop (using PowerShell or Command Prompt)
# Replace YOUR_EC2_IP and your-key.pem with actual values

scp -i your-key.pem ec2-user@YOUR_EC2_IP:/home/ec2-user/excel-ollama-deploy/ExcelOllamaPlugin-Windows-v1.0.0.zip .
```

**Alternative download methods:**
- Use **WinSCP** (GUI tool for Windows)
- Use **FileZilla** with SFTP
- Use **AWS S3** (upload from EC2, download to Windows)

## Part 2: Windows Installation (5 minutes)

### Step 1: Extract and Install

1. **Extract the ZIP file** you downloaded
2. **Open the extracted folder**
3. **Double-click `install-plugin.bat`**
4. **Follow the prompts** (it will install Python packages automatically)
5. **Wait for completion** (usually 2-3 minutes)

### Step 2: Verify Installation

1. **Restart Excel completely**
2. **Look for "Ollama AI Analysis" tab** in the ribbon
3. **If tab doesn't appear:**
   - Go to File → Options → Add-ins
   - Ensure xlwings add-in is enabled
   - Restart Excel again

### Step 3: Configure Connection

1. **Click "Configure"** in the Ollama AI Analysis tab
2. **Server URL should be pre-filled** with your EC2 IP
3. **Click "Test Connection"** - should show "Connected ✓"
4. **Select your preferred model** (llama2:latest recommended)
5. **Click "OK" to save**

## Part 3: Usage Examples (2 minutes)

### Basic Data Analysis

1. **Create or open Excel file** with data
2. **Select your data range** (including headers)
3. **Click "Analyze Data"** in Ollama AI Analysis tab
4. **Wait for analysis** (progress will show)
5. **Review results** in new "AI_Analysis_Results" sheet

### Natural Language Queries

1. **Select your data**
2. **Click "Ask Question"**
3. **Type questions like:**
   ```
   "What trends do you see in sales data?"
   "Are there any outliers in this dataset?"
   "Can you forecast the next 10 periods?"
   "What patterns exist in customer behavior?"
   ```
4. **Get AI-powered insights** in natural language

### Generate Reports

1. **Select your data**
2. **Click "Generate Report"**
3. **Get comprehensive analysis** with:
   - Executive summary
   - Key findings
   - Recommendations
   - Visualizations suggestions

## 🔧 Two Plugin Approaches

### Approach 1: Python-based Plugin (Default)

**Pros:**
- Full feature set
- Advanced analytics capabilities
- Extensible and customizable

**Cons:**
- Requires Python runtime on Windows
- Larger installation size

**Installation:** Use the `install-plugin.bat` from the EC2 build

### Approach 2: Web-based Excel Add-in (Alternative)

**Pros:**
- No Python required on Windows
- Lighter weight
- Uses Office.js (Microsoft's standard)

**Cons:**
- Limited to basic analysis features
- Requires hosting the web files

**Files for web approach:**
- `web-addin-manifest.xml`
- `web-addin-commands.html`
- `web-based-excel-addin.js`

## 🚨 Troubleshooting

### Common Issues

#### 1. "Cannot connect to Ollama"
```bash
# Check EC2 instance status
aws ec2 describe-instances --instance-ids i-YOUR_INSTANCE_ID

# SSH to EC2 and check Ollama
ssh -i your-key.pem ec2-user@YOUR_EC2_IP
sudo systemctl status ollama
curl localhost:11434/api/tags
```

#### 2. "Excel ribbon tab not showing"
- Restart Excel completely
- Check Excel Add-ins (File → Options → Add-ins)
- Enable xlwings add-in if disabled
- Run as Administrator if needed

#### 3. "Python not found" during installation
- Install Python 3.8+ from python.org
- **Important:** Check "Add Python to PATH" during installation
- Restart Command Prompt and try again

#### 4. "Connection timeout"
- Verify EC2 Security Group allows port 11434 from your IP
- Check if your public IP changed (whatismyipaddress.com)
- Test connection: `curl http://YOUR_EC2_IP:11434/api/tags`

### Testing Your Setup

Run these tests to verify everything works:

```bash
# Test 1: EC2 Ollama server (from Windows)
curl http://YOUR_EC2_IP:11434/api/tags

# Test 2: Excel plugin (in Excel)
# Select data → Click "Analyze Data" → Check for results

# Test 3: Natural language query
# Select data → Click "Ask Question" → Type "What trends do you see?"
```

## 💰 Cost Management

### EC2 Costs (US East):
- **t3.medium:** ~$1/day (24/7) or ~$0.30/day (8 hours)
- **t3.large:** ~$2/day (24/7) or ~$0.60/day (8 hours)

### Save Money:
```bash
# Stop EC2 when not analyzing (saves ~70% cost)
aws ec2 stop-instances --instance-ids i-YOUR_INSTANCE_ID

# Start when needed
aws ec2 start-instances --instance-ids i-YOUR_INSTANCE_ID

# Note: Ollama models stay loaded in memory when you restart
```

### Auto Start/Stop Script:
```bash
# Create a scheduled task on Windows to start/stop EC2
# Or use AWS Lambda with CloudWatch Events
```

## 🔒 Security Best Practices

1. **Restrict EC2 Access:**
   ```bash
   # In Security Group, only allow your IP:
   Type: Custom TCP
   Port: 11434
   Source: YOUR_LAPTOP_IP/32  # Not 0.0.0.0/0!
   ```

2. **Keep EC2 Updated:**
   ```bash
   sudo yum update -y  # Run monthly
   ```

3. **Monitor Usage:**
   - Set up CloudWatch alarms
   - Monitor costs in AWS Billing

## 📊 Performance Tips

### For Large Datasets:
- Use t3.large or t3.xlarge EC2 instances
- Enable data chunking in plugin settings
- Use data sampling for initial exploration

### For Faster Responses:
- Keep EC2 instance running during work sessions
- Use smaller models (phi, mistral) for quick queries
- Use larger models (llama2:13b) for complex analysis

## 🎉 Success Checklist

After following this guide, you should have:

- [ ] ✅ EC2 instance running Ollama with AI models
- [ ] ✅ Excel plugin installed on Windows laptop
- [ ] ✅ "Ollama AI Analysis" tab visible in Excel
- [ ] ✅ Successful connection test to EC2 server
- [ ] ✅ Ability to analyze data and ask questions
- [ ] ✅ Generated at least one analysis report

## 🆘 Getting Help

### Before asking for help, collect:
1. **EC2 instance details** (type, region, public IP)
2. **Windows version** and Excel version
3. **Error messages** from installation or usage
4. **Output of connection test:** `curl http://YOUR_EC2_IP:11434/api/tags`
5. **Plugin logs** (if available): `%APPDATA%\ExcelOllamaPlugin\logs\`

### Resources:
- **EC2 Build Script:** `ec2-build-and-deploy.sh`
- **Connection Test:** `test-connection.py`
- **AWS Documentation:** https://docs.aws.amazon.com/ec2/
- **Ollama Documentation:** https://ollama.ai/docs

---

**🎊 Congratulations!** You now have a powerful, cloud-based AI analysis system that works seamlessly with Excel, without requiring any Python development environment on your Windows laptop!

**Happy analyzing!** 📊✨