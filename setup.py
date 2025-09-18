"""
Setup script for Excel-Ollama AI Plugin.
"""

from setuptools import setup, find_packages
import os

# Read version from __init__.py
def get_version():
    with open(os.path.join('src', 'core', '__init__.py'), 'r') as f:
        for line in f:
            if line.startswith('__version__'):
                return line.split('=')[1].strip().strip('"\'')
    return '1.0.0'

# Read long description from README
def get_long_description():
    if os.path.exists('README.md'):
        with open('README.md', 'r', encoding='utf-8') as f:
            return f.read()
    return "Excel-Ollama AI Plugin for intelligent data analysis"

setup(
    name="excel-ollama-ai-plugin",
    version=get_version(),
    author="Excel-Ollama AI Plugin Team",
    author_email="support@excel-ollama-plugin.com",
    description="AI-powered Excel plugin using Ollama models for intelligent data analysis",
    long_description=get_long_description(),
    long_description_content_type="text/markdown",
    url="https://github.com/excel-ollama-ai-plugin/excel-ollama-ai-plugin",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Scientific/Engineering :: Artificial Intelligence",
    ],
    python_requires=">=3.8",
    install_requires=[
        "xlwings>=0.30.0",
        "openpyxl>=3.1.0",
        "requests>=2.31.0",
        "aiohttp>=3.8.0",
        "pandas>=2.0.0",
        "numpy>=1.24.0",
        "plotly>=5.15.0",
        "matplotlib>=3.7.0",
        "python-dateutil>=2.8.0",
        "pydantic>=2.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.4.0",
            "pytest-asyncio>=0.21.0",
            "pytest-mock>=3.11.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.5.0",
        ],
        "build": [
            "pyinstaller>=5.13.0",
            "setuptools>=68.0.0",
            "wheel>=0.41.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "excel-ollama-plugin=core.main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "ui": ["*.xml", "*.json"],
        "": ["*.md", "*.txt"],
    },
    zip_safe=False,
    keywords="excel plugin ai ollama data analysis machine learning",
    project_urls={
        "Bug Reports": "https://github.com/excel-ollama-ai-plugin/excel-ollama-ai-plugin/issues",
        "Source": "https://github.com/excel-ollama-ai-plugin/excel-ollama-ai-plugin",
        "Documentation": "https://excel-ollama-ai-plugin.readthedocs.io/",
    },
)