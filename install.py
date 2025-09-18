"""
Installation script for Excel-Ollama AI Plugin.
Handles plugin installation, registration, and setup.
"""

import os
import sys
import shutil
import subprocess
import winreg
import json
from pathlib import Path
import argparse
import logging


class PluginInstaller:
    """Handles installation and deployment of the Excel-Ollama AI Plugin."""
    
    def __init__(self):
        self.plugin_name = "ExcelOllamaAIPlugin"
        self.plugin_version = "1.0.0"
        self.setup_logging()
        
        # Paths
        self.source_dir = Path(__file__).parent
        self.install_dir = Path.home() / "AppData" / "Local" / self.plugin_name
        self.excel_addins_dir = Path.home() / "AppData" / "Roaming" / "Microsoft" / "Excel" / "XLSTART"
        
    def setup_logging(self):
        """Setup logging for installation process."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler('install.log')
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def check_prerequisites(self):
        """Check if all prerequisites are installed."""
        self.logger.info("Checking prerequisites...")
        
        # Check Python version
        if sys.version_info < (3, 8):
            raise Exception("Python 3.8 or later is required")
        
        # Check if Excel is installed
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                               r"SOFTWARE\Microsoft\Office")
            winreg.CloseKey(key)
            self.logger.info("Microsoft Office found")
        except:
            self.logger.warning("Microsoft Office not found in registry")
        
        # Check required Python packages
        required_packages = [
            'xlwings', 'pandas', 'numpy', 'requests', 'scikit-learn', 
            'scipy', 'aiohttp', 'tkinter'
        ]
        
        missing_packages = []
        for package in required_packages:
            try:
                __import__(package)
            except ImportError:
                missing_packages.append(package)
        
        if missing_packages:
            self.logger.info(f"Installing missing packages: {missing_packages}")
            self.install_packages(missing_packages)
        
        # Check if Ollama is available
        try:
            result = subprocess.run(['ollama', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                self.logger.info(f"Ollama found: {result.stdout.strip()}")
            else:
                self.logger.warning("Ollama not found. Please install Ollama separately.")
        except:
            self.logger.warning("Ollama not found. Please install Ollama separately.")
        
        self.logger.info("Prerequisites check completed")
    
    def install_packages(self, packages):
        """Install required Python packages."""
        for package in packages:
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                self.logger.info(f"Installed {package}")
            except subprocess.CalledProcessError as e:
                self.logger.error(f"Failed to install {package}: {e}")
                raise
    
    def create_directories(self):
        """Create necessary directories."""
        self.logger.info("Creating directories...")
        
        directories = [
            self.install_dir,
            self.install_dir / "src",
            self.install_dir / "logs",
            self.install_dir / "cache",
            self.install_dir / "config"
        ]
        
        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)
            self.logger.info(f"Created directory: {directory}")
    
    def copy_files(self):
        """Copy plugin files to installation directory."""
        self.logger.info("Copying plugin files...")
        
        # Copy source files
        src_source = self.source_dir / "src"
        src_dest = self.install_dir / "src"
        
        if src_source.exists():
            shutil.copytree(src_source, src_dest, dirs_exist_ok=True)
            self.logger.info("Copied source files")
        
        # Copy manifest
        manifest_source = self.source_dir / "manifest.xml"
        manifest_dest = self.install_dir / "manifest.xml"
        
        if manifest_source.exists():
            shutil.copy2(manifest_source, manifest_dest)
            self.logger.info("Copied manifest file")
        
        # Copy requirements
        req_source = self.source_dir / "requirements.txt"
        req_dest = self.install_dir / "requirements.txt"
        
        if req_source.exists():
            shutil.copy2(req_source, req_dest)
            self.logger.info("Copied requirements file")
        
        # Copy setup script
        setup_source = self.source_dir / "setup.py"
        setup_dest = self.install_dir / "setup.py"
        
        if setup_source.exists():
            shutil.copy2(setup_source, setup_dest)
            self.logger.info("Copied setup file")
    
    def register_excel_addin(self):
        """Register the plugin as an Excel add-in."""
        self.logger.info("Registering Excel add-in...")
        
        try:
            # Create xlwings add-in file
            addin_content = f'''
import sys
import os

# Add plugin path to Python path
plugin_path = r"{self.install_dir}"
if plugin_path not in sys.path:
    sys.path.insert(0, plugin_path)

# Import and initialize plugin
try:
    from src.main import initialize_plugin
    plugin = initialize_plugin()
    print("Excel-Ollama AI Plugin loaded successfully")
except Exception as e:
    print(f"Failed to load plugin: {{e}}")
    import traceback
    traceback.print_exc()
'''
            
            # Write add-in file
            addin_file = self.excel_addins_dir / f"{self.plugin_name}.py"
            self.excel_addins_dir.mkdir(parents=True, exist_ok=True)
            
            with open(addin_file, 'w') as f:
                f.write(addin_content)
            
            self.logger.info(f"Created Excel add-in file: {addin_file}")
            
            # Register with xlwings
            try:
                import xlwings as xw
                xw.Book(str(addin_file)).save()
                self.logger.info("Registered with xlwings")
            except Exception as e:
                self.logger.warning(f"xlwings registration failed: {e}")
            
        except Exception as e:
            self.logger.error(f"Failed to register Excel add-in: {e}")
            raise
    
    def create_configuration(self):
        """Create default configuration files."""
        self.logger.info("Creating configuration files...")
        
        config = {
            "ollama": {
                "server_url": "http://localhost:11434",
                "default_model": "llama2:latest",
                "timeout": 300,
                "max_retries": 3,
                "stream_responses": True
            },
            "analysis": {
                "default_type": "statistical_analysis",
                "auto_detect_types": True,
                "max_rows": 100000,
                "chunk_size": 10000,
                "parallel_processing": True
            },
            "ui": {
                "show_progress": True,
                "auto_open_results": True,
                "show_confidence": True,
                "notify_completion": True,
                "notify_errors": True
            },
            "advanced": {
                "log_level": "INFO",
                "enable_logging": True,
                "encrypt_cache": True,
                "clear_cache_on_exit": False,
                "cache_size_mb": 100
            }
        }
        
        config_file = self.install_dir / "config" / "config.json"
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=2)
        
        self.logger.info(f"Created configuration file: {config_file}")
    
    def register_windows_registry(self):
        """Register plugin in Windows registry."""
        self.logger.info("Registering in Windows registry...")
        
        try:
            # Create registry entries for uninstallation
            reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ExcelOllamaAIPlugin"
            
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, 
                                "Excel-Ollama AI Plugin")
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, 
                                self.plugin_version)
                winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, 
                                "Excel-Ollama AI Plugin")
                winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, 
                                str(self.install_dir))
                winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, 
                                f'python "{self.install_dir / "install.py"}" --uninstall')
                winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
            
            self.logger.info("Registry entries created")
            
        except Exception as e:
            self.logger.warning(f"Registry registration failed: {e}")
    
    def create_shortcuts(self):
        """Create desktop and start menu shortcuts."""
        self.logger.info("Creating shortcuts...")
        
        try:
            import win32com.client
            
            shell = win32com.client.Dispatch("WScript.Shell")
            
            # Desktop shortcut
            desktop = shell.SpecialFolders("Desktop")
            shortcut_path = os.path.join(desktop, f"{self.plugin_name}.lnk")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = sys.executable
            shortcut.Arguments = f'"{self.install_dir / "src" / "main.py"}"'
            shortcut.WorkingDirectory = str(self.install_dir)
            shortcut.IconLocation = sys.executable
            shortcut.save()
            
            self.logger.info("Desktop shortcut created")
            
        except ImportError:
            self.logger.warning("pywin32 not available, skipping shortcut creation")
        except Exception as e:
            self.logger.warning(f"Shortcut creation failed: {e}")
    
    def install(self):
        """Perform complete installation."""
        try:
            self.logger.info(f"Starting installation of {self.plugin_name} v{self.plugin_version}")
            
            self.check_prerequisites()
            self.create_directories()
            self.copy_files()
            self.create_configuration()
            self.register_excel_addin()
            self.register_windows_registry()
            self.create_shortcuts()
            
            self.logger.info("Installation completed successfully!")
            self.logger.info(f"Plugin installed to: {self.install_dir}")
            self.logger.info("Please restart Excel to use the plugin.")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Installation failed: {e}")
            return False
    
    def uninstall(self):
        """Uninstall the plugin."""
        try:
            self.logger.info(f"Starting uninstallation of {self.plugin_name}")
            
            # Remove Excel add-in
            addin_file = self.excel_addins_dir / f"{self.plugin_name}.py"
            if addin_file.exists():
                addin_file.unlink()
                self.logger.info("Removed Excel add-in file")
            
            # Remove installation directory
            if self.install_dir.exists():
                shutil.rmtree(self.install_dir)
                self.logger.info("Removed installation directory")
            
            # Remove registry entries
            try:
                reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ExcelOllamaAIPlugin"
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, reg_path)
                self.logger.info("Removed registry entries")
            except:
                pass
            
            # Remove shortcuts
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                desktop = shell.SpecialFolders("Desktop")
                shortcut_path = os.path.join(desktop, f"{self.plugin_name}.lnk")
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    self.logger.info("Removed desktop shortcut")
            except:
                pass
            
            self.logger.info("Uninstallation completed successfully!")
            return True
            
        except Exception as e:
            self.logger.error(f"Uninstallation failed: {e}")
            return False
    
    def update(self):
        """Update the plugin to latest version."""
        try:
            self.logger.info("Starting plugin update...")
            
            # Backup current configuration
            config_file = self.install_dir / "config" / "config.json"
            backup_config = None
            
            if config_file.exists():
                with open(config_file, 'r') as f:
                    backup_config = json.load(f)
                self.logger.info("Backed up current configuration")
            
            # Perform installation (will overwrite files)
            success = self.install()
            
            if success and backup_config:
                # Restore configuration
                with open(config_file, 'w') as f:
                    json.dump(backup_config, f, indent=2)
                self.logger.info("Restored configuration")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Update failed: {e}")
            return False


def main():
    """Main installation script entry point."""
    parser = argparse.ArgumentParser(description="Excel-Ollama AI Plugin Installer")
    parser.add_argument('--install', action='store_true', help='Install the plugin')
    parser.add_argument('--uninstall', action='store_true', help='Uninstall the plugin')
    parser.add_argument('--update', action='store_true', help='Update the plugin')
    parser.add_argument('--silent', action='store_true', help='Silent installation')
    
    args = parser.parse_args()
    
    installer = PluginInstaller()
    
    try:
        if args.uninstall:
            success = installer.uninstall()
        elif args.update:
            success = installer.update()
        else:
            success = installer.install()
        
        if success:
            if not args.silent:
                input("Press Enter to continue...")
            sys.exit(0)
        else:
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\nInstallation cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"Installation error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()