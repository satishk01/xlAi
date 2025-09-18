#!/usr/bin/env python3
"""
Windows Plugin Test Script
Tests the Excel-Ollama AI Plugin functionality on Windows
"""

import sys
import os
import json
import requests
from datetime import datetime

def test_python_environment():
    """Test Python environment and required packages."""
    print("Testing Python Environment")
    print("-" * 40)
    
    # Check Python version
    print(f"Python Version: {sys.version}")
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ required")
        return False
    else:
        print("✅ Python version OK")
    
    # Test required packages
    required_packages = [
        'xlwings', 'pandas', 'numpy', 'requests', 
        'aiohttp', 'scikit-learn', 'scipy'
    ]
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package}")
        except ImportError:
            print(f"❌ {package} - MISSING")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n❌ Missing packages: {', '.join(missing_packages)}")
        print("Run: pip install " + " ".join(missing_packages))
        return False
    
    return True

def test_plugin_installation():
    """Test if plugin is properly installed."""
    print("\nTesting Plugin Installation")
    print("-" * 40)
    
    # Check plugin directory
    plugin_dir = os.path.join(os.getenv('APPDATA'), 'ExcelOllamaPlugin')
    if os.path.exists(plugin_dir):
        print(f"✅ Plugin directory: {plugin_dir}")
    else:
        print(f"❌ Plugin directory not found: {plugin_dir}")
        return False
    
    # Check plugin files
    required_files = [
        'plugin/main.py',
        'plugin/core/ollama_client.py',
        'plugin/agents/analysis_agent.py'
    ]
    
    for file_path in required_files:
        full_path = os.path.join(plugin_dir, file_path)
        if os.path.exists(full_path):
            print(f"✅ {file_path}")
        else:
            print(f"❌ {file_path} - MISSING")
            return False
    
    # Check Excel startup file
    excel_startup = os.path.join(
        os.getenv('APPDATA'), 
        'Microsoft', 'Excel', 'XLSTART', 
        'ExcelOllamaPlugin.py'
    )
    
    if os.path.exists(excel_startup):
        print(f"✅ Excel startup file: {excel_startup}")
    else:
        print(f"❌ Excel startup file not found: {excel_startup}")
        print("   Plugin may not load automatically in Excel")
    
    return True

def test_configuration():
    """Test plugin configuration."""
    print("\nTesting Configuration")
    print("-" * 40)
    
    config_file = os.path.join(os.getenv('APPDATA'), 'ExcelOllamaPlugin', 'config.json')
    
    if os.path.exists(config_file):
        print(f"✅ Configuration file: {config_file}")
        
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
            
            server_url = config.get('ollama', {}).get('server_url', 'Not configured')
            default_model = config.get('ollama', {}).get('default_model', 'Not configured')
            
            print(f"✅ Server URL: {server_url}")
            print(f"✅ Default Model: {default_model}")
            
            return config
            
        except Exception as e:
            print(f"❌ Error reading config: {e}")
            return None
    else:
        print(f"❌ Configuration file not found: {config_file}")
        return None

def test_ollama_connection(config):
    """Test connection to Ollama server."""
    print("\nTesting Ollama Connection")
    print("-" * 40)
    
    if not config:
        print("❌ No configuration available")
        return False
    
    server_url = config.get('ollama', {}).get('server_url')
    if not server_url:
        print("❌ No server URL configured")
        return False
    
    print(f"Testing connection to: {server_url}")
    
    try:
        # Test basic connectivity
        response = requests.get(f"{server_url}/api/tags", timeout=10)
        
        if response.status_code == 200:
            print("✅ Connection successful!")
            
            data = response.json()
            models = data.get('models', [])
            
            print(f"✅ Available models: {len(models)}")
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
        print("❌ Connection timeout")
        print("   Check if EC2 instance is running")
        return False
    except requests.exceptions.ConnectionError:
        print("❌ Connection failed")
        print("   Check server URL and network connectivity")
        return False
    except Exception as e:
        print(f"❌ Connection error: {e}")
        return False

def test_excel_integration():
    """Test Excel integration."""
    print("\nTesting Excel Integration")
    print("-" * 40)
    
    try:
        import xlwings as xw
        print("✅ xlwings imported successfully")
        
        # Try to get Excel app (don't create new instance)
        try:
            apps = xw.apps
            if len(apps) > 0:
                print("✅ Excel application found")
            else:
                print("ℹ️  No Excel instances running (this is OK)")
        except Exception as e:
            print(f"ℹ️  Excel test skipped: {e}")
        
        return True
        
    except ImportError:
        print("❌ xlwings not available")
        return False
    except Exception as e:
        print(f"❌ Excel integration error: {e}")
        return False

def test_plugin_import():
    """Test importing plugin modules."""
    print("\nTesting Plugin Import")
    print("-" * 40)
    
    # Add plugin directory to path
    plugin_dir = os.path.join(os.getenv('APPDATA'), 'ExcelOllamaPlugin')
    if plugin_dir not in sys.path:
        sys.path.insert(0, plugin_dir)
    
    try:
        # Test importing main plugin modules
        from plugin.core.ollama_client import OllamaClient
        print("✅ OllamaClient imported")
        
        from plugin.agents.analysis_agent import AnalysisAgent
        print("✅ AnalysisAgent imported")
        
        from plugin.core.data_processor import DataProcessor
        print("✅ DataProcessor imported")
        
        from plugin.utils.config import PluginConfig
        print("✅ PluginConfig imported")
        
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Plugin import error: {e}")
        return False

def generate_test_report(results):
    """Generate test report."""
    print("\n" + "=" * 50)
    print("TEST REPORT")
    print("=" * 50)
    
    total_tests = len(results)
    passed_tests = sum(1 for result in results.values() if result)
    
    print(f"Total Tests: {total_tests}")
    print(f"Passed: {passed_tests}")
    print(f"Failed: {total_tests - passed_tests}")
    print(f"Success Rate: {(passed_tests/total_tests)*100:.1f}%")
    
    print("\nTest Results:")
    for test_name, result in results.items():
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"  {status} {test_name}")
    
    if passed_tests == total_tests:
        print("\n🎉 ALL TESTS PASSED!")
        print("Your Excel-Ollama AI Plugin is ready to use!")
        print("\nNext steps:")
        print("1. Open Microsoft Excel")
        print("2. Look for 'Ollama AI Analysis' tab in the ribbon")
        print("3. If tab doesn't appear, restart Excel")
        print("4. Click 'Configure' to verify settings")
        print("5. Start analyzing your data!")
    else:
        print("\n⚠️  SOME TESTS FAILED")
        print("Please address the failed tests before using the plugin.")
        print("Check the error messages above for guidance.")
    
    print("\nTest completed at:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

def main():
    """Main test function."""
    print("Excel-Ollama AI Plugin - Windows Test Suite")
    print("=" * 50)
    print("This script tests your plugin installation and configuration.")
    print("Make sure Excel is closed before running this test.")
    print()
    
    # Run all tests
    results = {}
    
    results["Python Environment"] = test_python_environment()
    results["Plugin Installation"] = test_plugin_installation()
    
    config = test_configuration()
    results["Configuration"] = config is not None
    
    results["Ollama Connection"] = test_ollama_connection(config)
    results["Excel Integration"] = test_excel_integration()
    results["Plugin Import"] = test_plugin_import()
    
    # Generate report
    generate_test_report(results)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user")
    except Exception as e:
        print(f"\n\nTest suite error: {e}")
        import traceback
        traceback.print_exc()
    
    input("\nPress Enter to exit...")