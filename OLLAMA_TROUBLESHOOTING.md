# 🔧 Ollama External Access Troubleshooting

## Issue: "No such file or directory" when accessing http://publicip:11434

This means Ollama is running but only listening on localhost, not accepting external connections.

## Quick Fix Steps:

### Step 1: Run the Fix Script (Recommended)

```bash
# SSH to your EC2 instance
ssh -i your-key.pem ec2-user@YOUR_EC2_PUBLIC_IP

# Download and run the fix script
curl -O https://raw.githubusercontent.com/your-repo/excel-ollama-plugin/main/fix-ollama-external-access.sh
chmod +x fix-ollama-external-access.sh
sudo ./fix-ollama-external-access.sh
```

### Step 2: Manual Fix (Alternative)

```bash
# SSH to EC2
ssh -i your-key.pem ec2-user@YOUR_EC2_PUBLIC_IP

# Stop Ollama
sudo systemctl stop ollama

# Create systemd override
sudo mkdir -p /etc/systemd/system/ollama.service.d/
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null << EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
Environment="OLLAMA_ORIGINS=*"
EOF

# Restart Ollama
sudo systemctl daemon-reload
sudo systemctl start ollama

# Wait and test
sleep 10
curl http://localhost:11434/api/tags
```

### Step 3: Verify External Access

From your Windows laptop, test:

```bash
# In Command Prompt or PowerShell
curl http://YOUR_EC2_PUBLIC_IP:11434/api/tags

# Or open in browser:
# http://YOUR_EC2_PUBLIC_IP:11434/api/tags
```

You should see JSON response with available models.

## Common Issues and Solutions:

### Issue 1: Still getting "connection refused"

**Check Security Group:**
1. Go to AWS Console → EC2 → Security Groups
2. Find your instance's security group
3. Add inbound rule:
   - Type: Custom TCP
   - Port: 11434
   - Source: Your laptop's public IP (find at whatismyipaddress.com)

### Issue 2: Ollama not starting

```bash
# Check logs
sudo journalctl -u ollama -f

# Check if port is in use
sudo netstat -tlnp | grep 11434

# Restart manually
sudo systemctl restart ollama
```

### Issue 3: Models not available

```bash
# Download models
ollama pull llama2:latest
ollama pull mistral:latest

# List models
ollama list
```

### Issue 4: Firewall blocking

```bash
# If using ufw firewall
sudo ufw allow 11434/tcp
sudo ufw reload

# Check firewall status
sudo ufw status
```

## Testing Commands:

### On EC2 (should work):
```bash
curl http://localhost:11434/api/tags
curl http://127.0.0.1:11434/api/tags
```

### From Windows (should work after fix):
```bash
curl http://YOUR_EC2_PUBLIC_IP:11434/api/tags
```

### Expected Response:
```json
{
  "models": [
    {
      "name": "llama2:latest",
      "modified_at": "2024-01-01T00:00:00Z",
      "size": 3825819519
    }
  ]
}
```

## Verification Checklist:

- [ ] ✅ Ollama service is running: `sudo systemctl status ollama`
- [ ] ✅ Listening on all interfaces: `sudo netstat -tlnp | grep 11434`
- [ ] ✅ Local access works: `curl localhost:11434/api/tags`
- [ ] ✅ External access works: `curl YOUR_EC2_IP:11434/api/tags`
- [ ] ✅ Security Group allows port 11434 from your IP
- [ ] ✅ Models are downloaded: `ollama list`

## Final Configuration:

After fixing, your Excel plugin should use:
- **Server URL:** `http://YOUR_EC2_PUBLIC_IP:11434`
- **Model:** `llama2:latest` (or your preferred model)

## Alternative: Use Private IP

If external access still doesn't work, you can:

1. **Use VPN** to connect to your VPC
2. **Use private IP** in Excel plugin: `http://PRIVATE_IP:11434`
3. **Set up SSH tunnel:**
   ```bash
   ssh -i your-key.pem -L 11434:localhost:11434 ec2-user@YOUR_EC2_IP
   # Then use http://localhost:11434 in Excel
   ```

The fix script should resolve the issue in most cases. Run it and test the external access!