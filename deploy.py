"""
Deployment script for Excel-Ollama AI Plugin.
Creates distributable package and handles deployment.
"""

import os
import sys
import shutil
import zipfile
import subprocess
from pathlib import Path
import json
from datetime import datetime


class PluginDeployer:
    """Handles plugin packaging and deployment."""
    
    def __init__(self):
        self.plugin_name = "ExcelOllamaAIPlugin"
        self.version = "1.0.0"
        self.source_dir = Path(__file__).parent
        self.build_dir = self.source_dir / "build"
        self.dist_dir = self.source_dir / "dist"
        
    def clean_build_dirs(self):
        """Clean build and dist directories."""
        print("Cleaning build directories...")
        
        for directory in [self.build_dir, self.dist_dir]:
            if directory.exists():
                shutil.rmtree(directory)
            directory.mkdir(parents=True, exist_ok=True)
    
    def copy_source_files(self):
        """Copy source files to build directory."""
        print("Copying source files...")
        
        # Files and directories to include
        include_items = [
            'src/',
            'manifest.xml',
            'requirements.txt',
            'setup.py',
            'install.py',
            'README.md',
            'LICENSE'
        ]
        
        for item in include_items:
            source_path = self.source_dir / item
            dest_path = self.build_dir / item
            
            if source_path.exists():
                if source_path.is_dir():
                    shutil.copytree(source_path, dest_path, dirs_exist_ok=True)
                else:
                    shutil.copy2(source_path, dest_path)
                print(f"Copied: {item}")
    
    def create_batch_installer(self):
        """Create Windows batch installer."""
        print("Creating batch installer...")
        
        batch_content = f'''@echo off
echo Installing {self.plugin_name} v{self.version}
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or later from https://python.org
    pause
    exit /b 1
)

REM Check if Ollama is installed
ollama --version >nul 2>&1
if errorlevel 1 (
    echo WARNING: Ollama is not installed
    echo Please install Ollama from https://ollama.ai
    echo You can continue installation and install Ollama later
    echo.
)

REM Install Python dependencies
echo Installing Python dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install Python dependencies
    pause
    exit /b 1
)

REM Run Python installer
echo Running plugin installer...
python install.py --install
if errorlevel 1 (
    echo ERROR: Plugin installation failed
    pause
    exit /b 1
)

echo.
echo Installation completed successfully!
echo Please restart Excel to use the plugin.
echo.
pause
'''
        
        batch_file = self.build_dir / "install.bat"
        with open(batch_file, 'w') as f:
            f.write(batch_content)
        
        print(f"Created: {batch_file}")
    
    def create_uninstaller(self):
        """Create uninstaller script."""
        print("Creating uninstaller...")
        
        uninstall_content = f'''@echo off
echo Uninstalling {self.plugin_name} v{self.version}
echo.

python install.py --uninstall
if errorlevel 1 (
    echo ERROR: Uninstallation failed
    pause
    exit /b 1
)

echo.
echo Uninstallation completed successfully!
echo.
pause
'''
        
        uninstall_file = self.build_dir / "uninstall.bat"
        with open(uninstall_file, 'w') as f:
            f.write(uninstall_content)
        
        print(f"Created: {uninstall_file}")
    
    def create_package_info(self):
        """Create package information file."""
        print("Creating package info...")
        
        package_info = {
            "name": self.plugin_name,
            "version": self.version,
            "description": "AI-powered data analysis plugin for Excel using Ollama models",
            "author": "Excel-Ollama AI Plugin Team",
            "license": "MIT",
            "build_date": datetime.now().isoformat(),
            "requirements": {
                "python": ">=3.8",
                "excel": ">=2016",
                "ollama": "latest"
            },
            "files": [
                "src/",
                "manifest.xml",
                "requirements.txt",
                "install.py",
                "install.bat",
                "uninstall.bat",
                "README.md"
            ]
        }
        
        info_file = self.build_dir / "package_info.json"
        with open(info_file, 'w') as f:
            json.dump(package_info, f, indent=2)
        
        print(f"Created: {info_file}")
    
    def create_zip_package(self):
        """Create ZIP package for distribution."""
        print("Creating ZIP package...")
        
        zip_filename = f"{self.plugin_name}_v{self.version}.zip"
        zip_path = self.dist_dir / zip_filename
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.build_dir):
                for file in files:
                    file_path = Path(root) / file
                    arc_path = file_path.relative_to(self.build_dir)
                    zipf.write(file_path, arc_path)
        
        print(f"Created ZIP package: {zip_path}")
        print(f"Package size: {zip_path.stat().st_size / 1024 / 1024:.1f} MB")
        
        return zip_path
    
    def create_installer_exe(self):
        """Create executable installer using PyInstaller (if available)."""
        print("Attempting to create executable installer...")
        
        try:
            # Check if PyInstaller is available
            subprocess.run(['pyinstaller', '--version'], 
                         capture_output=True, check=True)
            
            # Create installer spec
            installer_script = self.build_dir / "installer_main.py"
            installer_content = '''
import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from install import PluginInstaller

def main():
    installer = PluginInstaller()
    success = installer.install()
    
    if success:
        input("Installation completed! Press Enter to exit...")
    else:
        input("Installation failed! Press Enter to exit...")

if __name__ == "__main__":
    main()
'''
            
            with open(installer_script, 'w') as f:
                f.write(installer_content)
            
            # Run PyInstaller
            cmd = [
                'pyinstaller',
                '--onefile',
                '--windowed',
                '--name', f'{self.plugin_name}_Installer',
                '--distpath', str(self.dist_dir),
                str(installer_script)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("Executable installer created successfully!")
            else:
                print(f"PyInstaller failed: {result.stderr}")
                
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("PyInstaller not available, skipping executable creation")
    
    def generate_documentation(self):
        """Generate deployment documentation."""
        print("Generating deployment documentation...")
        
        doc_content = f'''# {self.plugin_name} v{self.version} - Deployment Package

## Contents

This package contains everything needed to install and use the Excel-Ollama AI Plugin.

### Files Included:
- `src/` - Plugin source code
- `install.bat` - Windows installer script
- `uninstall.bat` - Uninstaller script
- `install.py` - Python installation script
- `requirements.txt` - Python dependencies
- `manifest.xml` - Excel add-in manifest
- `README.md` - User documentation
- `package_info.json` - Package metadata

## Installation Instructions

### Prerequisites:
1. Microsoft Excel 2016 or later
2. Python 3.8 or later
3. Ollama (install from https://ollama.ai)

### Quick Installation:
1. Extract this package to a folder
2. Double-click `install.bat`
3. Follow the prompts
4. Restart Excel

### Manual Installation:
1. Extract the package
2. Open command prompt in the package folder
3. Run: `python install.py --install`
4. Restart Excel

## Usage

After installation:
1. Open Excel
2. Look for the "Ollama AI Analysis" tab in the ribbon
3. Click "Configure" to set up Ollama connection
4. Select your data and click "Analyze Data"

## Uninstallation

To remove the plugin:
1. Double-click `uninstall.bat`, or
2. Run: `python install.py --uninstall`

## Support

For help and documentation, see README.md or use the Help button in Excel.

Build Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
'''
        
        doc_file = self.build_dir / "DEPLOYMENT.md"
        with open(doc_file, 'w') as f:
            f.write(doc_content)
        
        print(f"Created: {doc_file}")
    
    def deploy(self):
        """Execute complete deployment process."""
        print(f"Starting deployment of {self.plugin_name} v{self.version}")
        print("=" * 60)
        
        try:
            self.clean_build_dirs()
            self.copy_source_files()
            self.create_batch_installer()
            self.create_uninstaller()
            self.create_package_info()
            self.generate_documentation()
            
            # Create distribution packages
            zip_package = self.create_zip_package()
            self.create_installer_exe()
            
            print("\n" + "=" * 60)
            print("Deployment completed successfully!")
            print(f"Distribution package: {zip_package}")
            print(f"Build directory: {self.build_dir}")
            print(f"Distribution directory: {self.dist_dir}")
            
            return True
            
        except Exception as e:
            print(f"Deployment failed: {e}")
            import traceback
            traceback.print_exc()
            return False


def main():
    """Main deployment function."""
    deployer = PluginDeployer()
    success = deployer.deploy()
    
    if success:
        print("\nDeployment package is ready for distribution!")
        print("Share the ZIP file or installer with users.")
    else:
        print("\nDeployment failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()