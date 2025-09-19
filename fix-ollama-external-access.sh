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
    
    # Test advanced models if they exist
    echo ""
    echo "Testing advanced AI models..."
    
    # Test Qwen 2.5
    if curl -s http://localhost:11434/api/tags | grep -q "qwen2.5:latest"; then
        echo "🧪 Testing Qwen 2.5 model..."
        QWEN_TEST=$(curl -s -X POST http://localhost:11434/api/generate \
            -H 'Content-Type: application/json' \
            -d '{"model":"qwen2.5:latest","prompt":"Hello","stream":false}' \
            --connect-timeout 30)
        
        if echo "$QWEN_TEST" | grep -q "response"; then
            echo "✅ Qwen 2.5 model working"
        else
            echo "❌ Qwen 2.5 model test failed"
        fi
    fi
    
    # Test DeepSeek thinking model
    if curl -s http://localhost:11434/api/tags | grep -q "deepseek-r1"; then
        echo "🧪 Testing DeepSeek thinking model..."
        DEEPSEEK_TEST=$(curl -s -X POST http://localhost:11434/api/generate \
            -H 'Content-Type: application/json' \
            -d '{"model":"deepseek-r1:latest","prompt":"Think about 2+2","stream":false}' \
            --connect-timeout 30)
        
        if echo "$DEEPSEEK_TEST" | grep -q "response"; then
            echo "✅ DeepSeek thinking model working"
        else
            echo "❌ DeepSeek thinking model test failed"
        fi
    fi
    
    # Test large Qwen model
    if curl -s http://localhost:11434/api/tags | grep -q "qwen2.5:32b"; then
        echo "🧪 Testing Qwen 2.5 32B (large model)..."
        echo "   Note: This may take longer due to model size..."
        QWEN_LARGE_TEST=$(curl -s -X POST http://localhost:11434/api/generate \
            -H 'Content-Type: application/json' \
            -d '{"model":"qwen2.5:32b","prompt":"Analyze","stream":false}' \
            --connect-timeout 60)
        
        if echo "$QWEN_LARGE_TEST" | grep -q "response"; then
            echo "✅ Qwen 2.5 32B model working"
        else
            echo "❌ Qwen 2.5 32B model test failed (may need more time/memory)"
        fi
    fi
    
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
    echo ""
    echo "Advanced AI Model Troubleshooting:"
    echo "=================================="
    echo "If advanced models are not working:"
    echo "1. Check available disk space: df -h"
    echo "2. Check memory usage: free -h"
    echo "3. Large models (32B) need 16GB+ RAM"
    echo "4. Restart Ollama: sudo systemctl restart ollama"
    echo "5. Check Ollama logs: sudo journalctl -u ollama -f"
    echo ""
    echo "Model-specific issues:"
    echo "• DeepSeek R1: Requires 8GB+ RAM for thinking"
    echo "• Qwen 2.5 32B: Requires 16GB+ RAM for optimal performance"
    echo "• If models fail to load, try smaller variants:"
    echo "  - qwen2.5:7b instead of qwen2.5:32b"
    echo "  - deepseek-r1:1.5b instead of deepseek-r1:latest"
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
echo "Checking for web server (port 3000):"
if sudo netstat -tlnp | grep :3000 || sudo ss -tlnp | grep :3000; then
    echo "✅ Web server detected on port 3000"
    
    # Test web server access
    if curl -s --connect-timeout 5 http://localhost:3000/health > /dev/null; then
        echo "✅ Web server health check passed"
        echo "🌐 Web interface available at: http://$EC2_PUBLIC_IP:3000"
    else
        echo "❌ Web server health check failed"
    fi
else
    echo "ℹ️  No web server detected on port 3000"
    echo "   If you want the web interface, run: ./web-server-setup.sh"
fi

echo ""
echo "Available models:"
MODELS=$(curl -s http://localhost:11434/api/tags | grep -o '"name":"[^"]*"' | cut -d'"' -f4)
if [ -n "$MODELS" ]; then
    echo "$MODELS"
    
    # Check for advanced AI models
    echo ""
    echo "Advanced AI Model Status:"
    echo "========================="
    
    if echo "$MODELS" | grep -q "qwen2.5"; then
        echo "✅ Qwen 2.5 models available"
    else
        echo "❌ Qwen 2.5 models missing"
    fi
    
    if echo "$MODELS" | grep -q "deepseek"; then
        echo "✅ DeepSeek thinking models available"
    else
        echo "❌ DeepSeek thinking models missing"
    fi
    
    if echo "$MODELS" | grep -q "llama2"; then
        echo "✅ Standard models (llama2) available"
    else
        echo "❌ Standard models missing"
    fi
    
    # Count total models
    MODEL_COUNT=$(echo "$MODELS" | wc -l)
    echo "📊 Total models installed: $MODEL_COUNT"
    
else
    echo "Could not fetch models"
fi

echo ""
echo "=========================================="
echo "ADVANCED AI MODEL INSTALLATION"
echo "=========================================="

# Check if advanced models are missing and offer to install them
MODELS=$(curl -s http://localhost:11434/api/tags | grep -o '"name":"[^"]*"' | cut -d'"' -f4)

MISSING_MODELS=()

if ! echo "$MODELS" | grep -q "qwen2.5:latest"; then
    MISSING_MODELS+=("qwen2.5:latest")
fi

if ! echo "$MODELS" | grep -q "qwen2.5:32b"; then
    MISSING_MODELS+=("qwen2.5:32b")
fi

if ! echo "$MODELS" | grep -q "deepseek-r1:latest"; then
    MISSING_MODELS+=("deepseek-r1:latest")
fi

if [ ${#MISSING_MODELS[@]} -gt 0 ]; then
    echo "🤖 Missing Advanced AI Models Detected!"
    echo ""
    echo "The following advanced models are not installed:"
    for model in "${MISSING_MODELS[@]}"; do
        echo "  ❌ $model"
    done
    echo ""
    echo "These models enable:"
    echo "  🧠 Advanced thinking and reasoning"
    echo "  🤖 GitHub Copilot-style analysis"
    echo "  📊 Enhanced data insights"
    echo "  🎯 Better question answering"
    echo ""
    
    read -p "Would you like to install the missing advanced models? (y/n): " -n 1 -r
    echo
    
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo ""
        echo "Installing advanced AI models..."
        echo "⚠️  This may take 20-45 minutes depending on your internet connection"
        echo ""
        
        for model in "${MISSING_MODELS[@]}"; do
            echo "📥 Installing $model..."
            if ollama pull "$model"; then
                echo "✅ Successfully installed $model"
            else
                echo "❌ Failed to install $model"
            fi
            echo ""
        done
        
        echo "🎉 Advanced model installation completed!"
        echo ""
        echo "Updated model list:"
        ollama list
    else
        echo "Skipping advanced model installation."
        echo "You can install them later with:"
        for model in "${MISSING_MODELS[@]}"; do
            echo "  ollama pull $model"
        done
    fi
else
    echo "✅ All advanced AI models are already installed!"
    echo ""
    echo "Available advanced features:"
    echo "  🧠 DeepSeek thinking models"
    echo "  🤖 Qwen 2.5 for Copilot analysis"
    echo "  📊 Enhanced data processing"
fi

echo ""
echo "=========================================="
echo "NEXT STEPS"
echo "=========================================="
echo "1. Test basic connection from your Windows laptop:"
echo "   curl http://$EC2_PUBLIC_IP:11434/api/tags"
echo ""
echo "2. Or open in browser:"
echo "   http://$EC2_PUBLIC_IP:11434/api/tags"
echo ""
echo "3. Test advanced model (if installed):"
echo "   curl -X POST http://$EC2_PUBLIC_IP:11434/api/generate \\"
echo "        -H 'Content-Type: application/json' \\"
echo "        -d '{\"model\":\"qwen2.5:latest\",\"prompt\":\"Hello\",\"stream\":false}'"
echo ""
echo "4. Update your Excel plugin configuration:"
echo "   📍 Server URL: http://$EC2_PUBLIC_IP:11434"
echo "   🤖 Default Model: qwen2.5:latest"
echo "   🧠 Thinking Model: deepseek-r1:latest"
echo "   🚀 Copilot Model: qwen2.5:32b"
echo ""
echo "5. If using web interface, access at:"
echo "   🌐 http://$EC2_PUBLIC_IP:3000"
echo ""
echo "6. If still not working, check EC2 Security Group:"
echo "   🔒 Port 11434 (Ollama API)"
echo "   🔒 Port 3000 (Web interface, if using)"
echo "   🔒 Protocol: TCP"
echo "   🔒 Source: Your laptop's public IP or 0.0.0.0/0"
echo ""
echo "7. Advanced Features Available:"
echo "   🧠 Thinking models for complex reasoning"
echo "   🤖 GitHub Copilot-style comprehensive analysis"
echo "   📊 Native Excel chart generation"
echo "   🎯 Enhanced question answering"
echo "=========================================="