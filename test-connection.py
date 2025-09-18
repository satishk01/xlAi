#!/usr/bin/env python3
"""
Connection Test Script for Excel-Ollama AI Plugin
Tests connection to remote Ollama server on EC2
"""

import requests
import json
import sys
import time
from urllib.parse import urlparse


def test_ollama_connection(server_url, timeout=10):
    """Test connection to Ollama server."""
    print(f"Testing connection to: {server_url}")
    print("-" * 50)
    
    try:
        # Parse URL to validate format
        parsed = urlparse(server_url)
        if not parsed.scheme or not parsed.netloc:
            print("❌ Invalid URL format. Use: http://your-ec2-ip:11434")
            return False
        
        # Test 1: Basic connectivity
        print("1. Testing basic connectivity...")
        response = requests.get(f"{server_url}/api/tags", timeout=timeout)
        
        if response.status_code == 200:
            print("✅ Connection successful!")
            
            # Parse response
            data = response.json()
            models = data.get('models', [])
            
            print(f"✅ Found {len(models)} models:")
            for model in models:
                name = model.get('name', 'Unknown')
                size = model.get('size', 0)
                size_mb = size / (1024 * 1024) if size else 0
                print(f"   - {name} ({size_mb:.1f} MB)")
            
            return True
        else:
            print(f"❌ HTTP Error: {response.status_code}")
            return False
            
    except requests.exceptions.ConnectTimeout:
        print("❌ Connection timeout - check if server is running")
        return False
    except requests.exceptions.ConnectionError:
        print("❌ Connection failed - check URL and network")
        return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Request failed: {e}")
        return False
    except json.JSONDecodeError:
        print("❌ Invalid response format")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False


def test_model_generation(server_url, model_name="llama2", timeout=30):
    """Test model generation capability."""
    print(f"\n2. Testing model generation with {model_name}...")
    
    try:
        payload = {
            "model": model_name,
            "prompt": "Hello, this is a test. Please respond with 'Test successful!'",
            "stream": False
        }
        
        response = requests.post(
            f"{server_url}/api/generate",
            json=payload,
            timeout=timeout
        )
        
        if response.status_code == 200:
            data = response.json()
            response_text = data.get('response', '').strip()
            
            if response_text:
                print("✅ Model generation successful!")
                print(f"   Response: {response_text[:100]}...")
                return True
            else:
                print("❌ Empty response from model")
                return False
        else:
            print(f"❌ Generation failed: HTTP {response.status_code}")
            return False
            
    except requests.exceptions.Timeout:
        print("❌ Generation timeout - model may be loading")
        return False
    except Exception as e:
        print(f"❌ Generation error: {e}")
        return False


def get_server_info(server_url):
    """Get server information."""
    print("\n3. Getting server information...")
    
    try:
        # Try to get version info (if available)
        response = requests.get(f"{server_url}/api/version", timeout=5)
        if response.status_code == 200:
            version_info = response.json()
            print(f"✅ Server version: {version_info}")
        
    except:
        print("ℹ️  Version info not available")
    
    try:
        # Get model list with details
        response = requests.get(f"{server_url}/api/tags", timeout=10)
        if response.status_code == 200:
            data = response.json()
            models = data.get('models', [])
            
            total_size = sum(model.get('size', 0) for model in models)
            total_size_gb = total_size / (1024 * 1024 * 1024)
            
            print(f"✅ Total models: {len(models)}")
            print(f"✅ Total size: {total_size_gb:.2f} GB")
            
    except Exception as e:
        print(f"ℹ️  Could not get detailed info: {e}")


def main():
    """Main test function."""
    print("Excel-Ollama AI Plugin - Connection Test")
    print("=" * 50)
    
    # Get server URL from command line or prompt user
    if len(sys.argv) > 1:
        server_url = sys.argv[1]
    else:
        print("Enter your EC2 Ollama server URL:")
        print("Example: http://3.15.123.45:11434")
        server_url = input("URL: ").strip()
    
    if not server_url:
        print("❌ No URL provided")
        sys.exit(1)
    
    # Ensure URL has proper format
    if not server_url.startswith(('http://', 'https://')):
        server_url = f"http://{server_url}"
    
    if not server_url.endswith(':11434') and ':' not in server_url.split('://', 1)[1]:
        server_url = f"{server_url}:11434"
    
    print(f"\nTesting server: {server_url}")
    print("=" * 50)
    
    # Run tests
    connection_ok = test_ollama_connection(server_url)
    
    if connection_ok:
        # Test model generation
        test_model_generation(server_url)
        
        # Get server info
        get_server_info(server_url)
        
        print("\n" + "=" * 50)
        print("✅ Connection test completed successfully!")
        print("✅ Your Excel plugin should work with this server.")
        print("\nNext steps:")
        print("1. Open Excel")
        print("2. Go to 'Ollama AI Analysis' tab")
        print("3. Click 'Configure'")
        print(f"4. Set Server URL to: {server_url}")
        print("5. Click 'Test Connection'")
        print("6. Start analyzing your data!")
        
    else:
        print("\n" + "=" * 50)
        print("❌ Connection test failed!")
        print("\nTroubleshooting steps:")
        print("1. Check if EC2 instance is running")
        print("2. Verify Security Group allows port 11434")
        print("3. Confirm Ollama service is running on EC2:")
        print("   ssh to EC2 and run: sudo systemctl status ollama")
        print("4. Test from EC2 itself: curl localhost:11434/api/tags")
        print("5. Check your laptop's public IP hasn't changed")
        
        sys.exit(1)


if __name__ == "__main__":
    main()