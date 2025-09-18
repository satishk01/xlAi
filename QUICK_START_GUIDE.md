# 🚀 Quick Start Guide: Excel-Ollama AI Plugin with EC2

## Overview
This guide will get you up and running in 30 minutes with:
- Ollama AI server on Amazon EC2
- Excel plugin on your laptop
- Full AI-powered data analysis capabilities

## 📋 Prerequisites Checklist

### On Your Laptop:
- [ ] Windows 10/11
- [ ] Microsoft Excel 2016 or later
- [ ] Python 3.8+ installed ([Download here](https://python.org))
- [ ] Git installed ([Download here](https://git-scm.com))
- [ ] AWS account with EC2 access

### Knowledge Required:
- [ ] Basic AWS EC2 usage
- [ ] Basic command line usage
- [ ] Excel basics

## 🎯 Step-by-Step Setup

### Phase 1: Set Up EC2 Server (15 minutes)

#### 1.1 Launch EC2 Instance
```bash
# In AWS Console:
# 1. Go to EC2 → Launch Instance
# 2. Choose: Amazon Linux 2 AMI
# 3. Instance Type: t3.medium (minimum) or t3.large (recommended)
# 4. Storage: 20 GB
# 5. Create new Key Pair (save the .pem file)
# 6. Launch instance
```

#### 1.2 Configure Security Group
```bash
# In AWS Console:
# 1. Go to Security Groups
# 2. Select your instance's security group
# 3. Add Inbound Rule:
#    - Type: Custom TCP
#    - Port: 11434
#    - Source: Your IP (find at whatismyipaddress.com)
```

#### 1.3 Install Ollama on EC2
```bash
# Connect to EC2
ssh -i your-key.pem ec2-user@YOUR_EC2_PUBLIC_IP

# Run the automated setup
curl -fsSL https://ollama.ai/install.sh | sh

# Configure for external access
sudo mkdir -p /etc/systemd/system/ollama.service.d/
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null <<EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
EOF

# Restart Ollama
sudo systemctl daemon-reload
sudo systemctl restart ollama
sudo systemctl enable ollama

# Download AI models (this takes 5-10 minutes)
ollama pull llama2:latest
ollama pull mistral:latest

# Test installation
curl localhost:11434/api/tags
```

#### 1.4 Get Your EC2 Details
```bash
# Note these for later:
echo "Public IP: $(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)"
echo "Your Ollama URL: http://$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4):11434"
```

### Phase 2: Set Up Excel Plugin (10 minutes)

#### 2.1 Download Plugin Code
```bash
# On your laptop, open Command Prompt or PowerShell
git clone https://github.com/your-repo/excel-ollama-plugin.git
cd excel-ollama-plugin

# Or download ZIP and extract
```

#### 2.2 Install Plugin
```bash
# Run the automated setup
laptop-setup.bat

# Or manual installation:
pip install -r requirements.txt
python install.py --install
```

#### 2.3 Test Connection
```bash
# Test connection to your EC2 server
python test-connection.py http://YOUR_EC2_PUBLIC_IP:11434
```

### Phase 3: Configure and Use (5 minutes)

#### 3.1 Configure Excel Plugin
1. **Open Microsoft Excel**
2. **Look for "Ollama AI Analysis" tab** in the ribbon
3. **Click "Configure"**
4. **Enter your settings:**
   ```
   Server URL: http://YOUR_EC2_PUBLIC_IP:11434
   Default Model: llama2:latest
   Timeout: 300
   ```
5. **Click "Test Connection"** - should show "Connected ✓"
6. **Click "OK" to save**

#### 3.2 First Analysis
1. **Create sample data** or open existing Excel file
2. **Select your data range** (including headers)
3. **Click "Analyze Data"** in the Ollama AI Analysis tab
4. **Wait for analysis** (progress bar will show)
5. **Review results** in the new sheet that opens

#### 3.3 Try Natural Language Queries
1. **Select your data**
2. **Click "Ask Question"**
3. **Type:** "What trends do you see in this data?"
4. **Click "Analyze"**
5. **Read the AI-generated insights**

## 🎉 You're Done!

Your Excel-Ollama AI Plugin is now fully operational! 

## 📊 What You Can Do Now

### Basic Analysis
- **Statistical Analysis:** Descriptive statistics, correlations
- **Trend Analysis:** Time series trends, forecasting
- **Pattern Detection:** Seasonal patterns, anomalies
- **Clustering:** Group similar data points

### Advanced Features
- **Natural Language Queries:** Ask questions in plain English
- **Custom Excel Functions:** Use `=OLLAMA_ANALYZE()` in cells
- **Automated Reports:** Generate comprehensive analysis reports
- **Interactive Dashboards:** Create executive dashboards

### Example Queries to Try
```
"What are the key trends in sales data?"
"Are there any outliers in this dataset?"
"Can you forecast the next 10 periods?"
"What patterns do you see in customer behavior?"
"Group this data into meaningful clusters"
```

## 🔧 Troubleshooting

### Common Issues

#### "Cannot connect to Ollama"
```bash
# Check EC2 instance status
aws ec2 describe-instances --instance-ids i-YOUR_INSTANCE_ID

# SSH to EC2 and check Ollama
ssh -i your-key.pem ec2-user@YOUR_EC2_IP
sudo systemctl status ollama

# Test locally on EC2
curl localhost:11434/api/tags
```

#### "Connection timeout"
- Verify Security Group allows port 11434 from your IP
- Check if your public IP changed (use whatismyipaddress.com)
- Try increasing timeout in Excel plugin configuration

#### "Model not found"
```bash
# SSH to EC2 and check models
ollama list

# Download missing models
ollama pull llama2:latest
```

#### Excel plugin not visible
- Restart Excel completely
- Check if xlwings add-in is enabled in Excel Options
- Verify plugin installation: `python -c "import xlwings as xw; print(xw.addin)"`

## 💰 Cost Management

### Typical Monthly Costs (US East):
- **t3.medium:** ~$30/month (24/7) or ~$9/month (8 hours/day)
- **t3.large:** ~$60/month (24/7) or ~$18/month (8 hours/day)

### Save Money:
```bash
# Stop EC2 when not in use (saves ~70% cost)
aws ec2 stop-instances --instance-ids i-YOUR_INSTANCE_ID

# Start when needed
aws ec2 start-instances --instance-ids i-YOUR_INSTANCE_ID
```

## 🔒 Security Best Practices

1. **Restrict Access:** Only allow your IP in Security Group
2. **Keep Updated:** Regularly update EC2 instance
3. **Monitor Usage:** Set up CloudWatch alarms
4. **Use IAM Roles:** Don't store AWS keys on EC2

## 📈 Performance Tips

### For Large Datasets:
- Enable chunking in plugin settings (default: 10,000 rows)
- Use data sampling for initial exploration
- Consider upgrading to t3.large or t3.xlarge

### For Faster Responses:
- Use smaller models (mistral vs llama2:13b)
- Keep models loaded (they stay in memory)
- Enable streaming responses

## 🆘 Getting Help

### Resources:
- **Full Documentation:** See `DEPLOYMENT_GUIDE.md`
- **Test Connection:** Run `python test-connection.py`
- **Plugin Logs:** Check `%APPDATA%\ExcelOllamaPlugin\logs\`
- **Ollama Docs:** https://ollama.ai/docs

### Support Checklist:
When asking for help, include:
- [ ] EC2 instance type and region
- [ ] Output of `python test-connection.py`
- [ ] Excel version and Windows version
- [ ] Error messages from plugin logs
- [ ] Screenshot of Excel ribbon (if UI issue)

---

**🎊 Congratulations!** You now have a powerful AI-enhanced Excel setup that can analyze data, detect patterns, and provide insights using state-of-the-art language models running on your own cloud infrastructure!

**Happy analyzing!** 📊✨