#!/bin/bash
# EC2 Ollama Setup Script for Excel-Ollama AI Plugin
# Run this script on your Amazon EC2 Linux instance

echo "Setting up Ollama on Amazon EC2 Linux..."
echo "=========================================="

# Update system
echo "Updating system packages..."
sudo yum update -y || sudo apt update -y

# Install required packages
echo "Installing required packages..."
if command -v yum &> /dev/null; then
    # Amazon Linux / RHEL / CentOS
    sudo yum install -y curl wget git python3 python3-pip
elif command -v apt &> /dev/null; then
    # Ubuntu / Debian
    sudo apt install -y curl wget git python3 python3-pip
fi

# Install Ollama
echo "Installing Ollama..."
curl -fsSL https://ollama.ai/install.sh | sh

# Start Ollama service
echo "Starting Ollama service..."
sudo systemctl enable ollama
sudo systemctl start ollama

# Wait for Ollama to start
echo "Waiting for Ollama to initialize..."
sleep 10

# Download recommended models
echo "Downloading AI models (this may take a while)..."
ollama pull llama2:latest
ollama pull codellama:latest
ollama pull mistral:latest

# Configure Ollama to accept external connections
echo "Configuring Ollama for external access..."
sudo mkdir -p /etc/systemd/system/ollama.service.d/
sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null <<EOF
[Service]
Environment="OLLAMA_HOST=0.0.0.0:11434"
EOF

# Restart Ollama with new configuration
sudo systemctl daemon-reload
sudo systemctl restart ollama

# Configure firewall (if ufw is available)
if command -v ufw &> /dev/null; then
    echo "Configuring firewall..."
    sudo ufw allow 11434/tcp
    sudo ufw --force enable
fi

# Get instance information
INSTANCE_IP=$(curl -s http://169.254.169.254/latest/meta-data/public-ipv4)
PRIVATE_IP=$(curl -s http://169.254.169.254/latest/meta-data/local-ipv4)

echo ""
echo "=========================================="
echo "Ollama Setup Complete!"
echo "=========================================="
echo "Public IP: $INSTANCE_IP"
echo "Private IP: $PRIVATE_IP"
echo "Ollama URL: http://$INSTANCE_IP:11434"
echo ""
echo "Available models:"
ollama list
echo ""
echo "Test the installation:"
echo "curl http://$INSTANCE_IP:11434/api/tags"
echo ""
echo "IMPORTANT: Make sure to configure your EC2 Security Group"
echo "to allow inbound traffic on port 11434 from your laptop's IP"
echo "=========================================="