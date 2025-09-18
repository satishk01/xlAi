# EC2 Security Group Configuration

## Setting up Security Group for Ollama Access

### Step 1: Create/Modify Security Group

1. **Go to AWS Console** → EC2 → Security Groups
2. **Select your instance's security group** or create a new one
3. **Add Inbound Rule:**
   - **Type:** Custom TCP
   - **Port Range:** 11434
   - **Source:** Your laptop's public IP address (find it at https://whatismyipaddress.com/)
   - **Description:** Ollama API access for Excel plugin

### Step 2: Alternative - Allow from Specific IP Range

If your IP changes frequently, you can:
- Use your ISP's IP range (less secure)
- Use a VPN with static IP
- Set up AWS VPN or Direct Connect

### Step 3: Security Best Practices

```bash
# On EC2 instance, you can also use iptables for additional security
sudo iptables -A INPUT -p tcp --dport 11434 -s YOUR_LAPTOP_IP -j ACCEPT
sudo iptables -A INPUT -p tcp --dport 11434 -j DROP

# Save iptables rules
sudo iptables-save > /etc/iptables/rules.v4
```

### Step 4: Test Connection

From your laptop, test the connection:
```bash
curl http://YOUR_EC2_PUBLIC_IP:11434/api/tags
```

You should see a JSON response with available models.

## Important Security Notes

⚠️ **Security Warning:** Only allow access from your specific IP address. Never use 0.0.0.0/0 for production use.

✅ **Recommended:** Use AWS VPC and private subnets with VPN access for production deployments.