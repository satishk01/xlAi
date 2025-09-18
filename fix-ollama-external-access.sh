#!/bin/bash
# Fix Ollama External Access on EC2
# This script configures Ollama to accept external connections

echo "Fixing Ollama External Access on EC2"
echo "===================================="

# Get current status
echo "Current Ollama status:"
sudo systemctl status ollama --no-pager

echo ""
echo "Checking current Ollama configuration..."

# Check if Ollama is running
if ! pgrep -f ollama > /dev/null; then
    echo "❌ Ollama is not running. Starting it..."
    sudo systemctl start ollama
    sleep 5
fi

# Get EC2 public IP
EC2_PUBLIC_IP=$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)
EC2_PRIVATE_IP=$(curl -s http://169.254.169.254/latest/meta-data/local-ipv4)

echo "EC2 Public IP: $EC2_PUBLIC_IP"
echo "EC2 Private IP: $EC2_PRIVATE_IP"

# Test current local access
echo ""
echo "Testing local access..."
if curl -s http://localhost:11434/api/tags > /dev/null; then
    echo "✅ Local access works"
else
    echo "❌ Local access failed"
fi

# Test current external access
echo ""
echo "Testing external access..."
if curl -s --connect-timeout 5 http://$EC2_PUBLIC_IP:11434/api/tags > /dev/null; then
    echo "✅ External access already works!"
    echo "Your Ollama URL: http://$EC2_PUBLIC_IP:11434"
    exit 0
else
    echo "❌ External access failed - fixing now..."
fi

# Stop Ollama service
echo ""
echo "Stopping Ollama service..."
sudo systemctl stop ollama

# Create systemd override directory
echo "Creating systemd override configuration..."
sudo mkdir -p /etc/systemd/system/ollama.service.d/

# Create override configuration to bind to all interfaces
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null << EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
Environment="OLLAMA_ORIGINS=*"
EOF

echo "✅ Created systemd override configuration"

# Reload systemd and restart Ollama
echo ""
echo "Reloading systemd and starting Ollama..."
sudo systemctl daemon-reload
sudo systemctl enable ollama
sudo systemctl start ollama

# Wait for Ollama to start
echo "Waiting for Ollama to start..."
sleep 10

# Check if Ollama is running
if sudo systemctl is-active --quiet ollama; then
    echo "✅ Ollama service is running"
else
    echo "❌ Ollama service failed to start"
    echo "Checking logs..."
    sudo journalctl -u ollama --no-pager -n 20
    exit 1
fi

# Test local access again
echo ""
echo "Testing local access..."
if curl -s http://localhost:11434/api/tags > /dev/null; then
    echo "✅ Local access works"
else
    echo "❌ Local access still failed"
    echo "Checking what's listening on port 11434..."
    sudo netstat -tlnp | grep 11434 || sudo ss -tlnp | grep 11434
fi

# Test external access
echo ""
echo "Testing external access..."
sleep 5  # Give it a moment

if curl -s --connect-timeout 10 http://$EC2_PUBLIC_IP:11434/api/tags > /dev/null; then
    echo "✅ External access now works!"
else
    echo "❌ External access still failed"
    echo ""
    echo "Checking network configuration..."
    
    # Check what's listening
    echo "Processes listening on port 11434:"
    sudo netstat -tlnp | grep 11434 || sudo ss -tlnp | grep 11434
    
    echo ""
    echo "Ollama process details:"
    ps aux | grep ollama | grep -v grep
    
    echo ""
    echo "Checking firewall (if ufw is installed)..."
    if command -v ufw &> /dev/null; then
        sudo ufw status
        echo "Adding firewall rule for port 11434..."
        sudo ufw allow 11434/tcp
    fi
    
    echo ""
    echo "Manual troubleshooting steps:"
    echo "1. Check EC2 Security Group allows port 11434"
    echo "2. Verify your laptop's public IP is allowed"
    echo "3. Try: curl http://$EC2_PRIVATE_IP:11434/api/tags (from another EC2 instance)"
fi

# Show final status
echo ""
echo "=========================================="
echo "FINAL STATUS"
echo "=========================================="
echo "Ollama service status:"
sudo systemctl status ollama --no-pager -l

echo ""
echo "Configuration:"
cat /etc/systemd/system/ollama.service.d/override.conf

echo ""
echo "Network listening:"
sudo netstat -tlnp | grep 11434 || sudo ss -tlnp | grep 11434

echo ""
echo "Available models:"
curl -s http://localhost:11434/api/tags | grep -o '"name":"[^"]*"' | cut -d'"' -f4 || echo "Could not fetch models"

echo ""
echo "=========================================="
echo "NEXT STEPS"
echo "=========================================="
echo "1. Test from your Windows laptop:"
echo "   curl http://$EC2_PUBLIC_IP:11434/api/tags"
echo ""
echo "2. Or open in browser:"
echo "   http://$EC2_PUBLIC_IP:11434/api/tags"
echo ""
echo "3. Update your Excel plugin with:"
echo "   Server URL: http://$EC2_PUBLIC_IP:11434"
echo ""
echo "4. If still not working, check EC2 Security Group:"
echo "   - Port: 11434"
echo "   - Protocol: TCP"
echo "   - Source: Your laptop's public IP"
echo "=========================================="