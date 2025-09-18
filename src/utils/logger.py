"""
Logging utilities for Excel-Ollama AI Plugin.
"""

import logging
import os
from datetime import datetime
from typing import Optional


def setup_logging(level: str = 'INFO', log_file: Optional[str] = None) -> logging.Logger:
    """Set up logging configuration for the plugin."""
    
    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ExcelOllamaPlugin', 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    # Default log file
    if log_file is None:
        timestamp = datetime.now().strftime('%Y%m%d')
        log_file = os.path.join(log_dir, f'plugin_{timestamp}.log')
    
    # Configure logging
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()  # Also log to console
        ]
    )
    
    logger = logging.getLogger('ExcelOllamaPlugin')
    logger.info(f"Logging initialized - Level: {level}, File: {log_file}")
    
    return logger


def get_logger(name: str) -> logging.Logger:
    """Get a logger instance for a specific module."""
    return logging.getLogger(f'ExcelOllamaPlugin.{name}')