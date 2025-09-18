"""
Configuration management for Excel-Ollama AI Plugin.
Handles plugin settings, Ollama configuration, and user preferences.
"""

import json
import os
from dataclasses import dataclass, asdict
from typing import Dict, Any, Optional
from pathlib import Path


@dataclass
class OllamaConfig:
    """Configuration for Ollama integration."""
    server_url: str = "http://localhost:11434"  # Can be EC2 instance IP
    default_model: str = "llama2"
    timeout: int = 300
    max_retries: int = 3
    stream_responses: bool = True
    model_parameters: Dict[str, Any] = None
    api_key: str = ""  # For future authentication support
    verify_ssl: bool = True  # SSL verification for HTTPS connections
    connection_test_timeout: int = 10  # Timeout for connection tests
    
    def __post_init__(self):
        if self.model_parameters is None:
            self.model_parameters = {
                "temperature": 0.7,
                "top_p": 0.9,
                "top_k": 40
            }


@dataclass
class ExcelSettings:
    """Excel-specific configuration settings."""
    auto_refresh: bool = True
    max_rows_per_chunk: int = 10000
    default_chart_type: str = "line"
    enable_custom_functions: bool = True
    ribbon_position: str = "right"


@dataclass
class AgentSettings:
    """Configuration for AI agents."""
    analysis_confidence_threshold: float = 0.7
    pattern_detection_sensitivity: float = 0.8
    max_concurrent_agents: int = 3
    agent_timeout: int = 120
    enable_agent_logging: bool = True


@dataclass
class UIPreferences:
    """User interface preferences."""
    theme: str = "light"
    show_progress_dialogs: bool = True
    auto_save_results: bool = True
    default_language: str = "en"
    show_tooltips: bool = True


@dataclass
class PluginConfig:
    """Main plugin configuration container."""
    ollama: OllamaConfig
    excel_settings: ExcelSettings
    agent_settings: AgentSettings
    ui_preferences: UIPreferences
    version: str = "1.0.0"
    
    def __post_init__(self):
        if not isinstance(self.ollama, OllamaConfig):
            self.ollama = OllamaConfig(**self.ollama)
        if not isinstance(self.excel_settings, ExcelSettings):
            self.excel_settings = ExcelSettings(**self.excel_settings)
        if not isinstance(self.agent_settings, AgentSettings):
            self.agent_settings = AgentSettings(**self.agent_settings)
        if not isinstance(self.ui_preferences, UIPreferences):
            self.ui_preferences = UIPreferences(**self.ui_preferences)


class ConfigManager:
    """Manages plugin configuration loading, saving, and validation."""
    
    def __init__(self, config_dir: Optional[str] = None):
        """Initialize configuration manager."""
        if config_dir is None:
            # Default to user's AppData directory
            self.config_dir = Path(os.getenv('APPDATA', '')) / 'ExcelOllamaPlugin'
        else:
            self.config_dir = Path(config_dir)
        
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.config_file = self.config_dir / 'config.json'
        self._config: Optional[PluginConfig] = None
    
    def load_config(self) -> PluginConfig:
        """Load configuration from file or create default."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                self._config = PluginConfig(**config_data)
            except (json.JSONDecodeError, TypeError, ValueError) as e:
                print(f"Error loading config: {e}. Using default configuration.")
                self._config = self._create_default_config()
        else:
            self._config = self._create_default_config()
            self.save_config()
        
        return self._config
    
    def save_config(self) -> bool:
        """Save current configuration to file."""
        if self._config is None:
            return False
        
        try:
            config_dict = asdict(self._config)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2, ensure_ascii=False)
            return True
        except (IOError, TypeError) as e:
            print(f"Error saving config: {e}")
            return False
    
    def get_config(self) -> PluginConfig:
        """Get current configuration, loading if necessary."""
        if self._config is None:
            return self.load_config()
        return self._config
    
    def update_config(self, **kwargs) -> bool:
        """Update configuration with new values."""
        if self._config is None:
            self.load_config()
        
        try:
            # Update nested configuration objects
            for key, value in kwargs.items():
                if hasattr(self._config, key):
                    if isinstance(value, dict):
                        # Update nested configuration
                        current_value = getattr(self._config, key)
                        if hasattr(current_value, '__dict__'):
                            for nested_key, nested_value in value.items():
                                if hasattr(current_value, nested_key):
                                    setattr(current_value, nested_key, nested_value)
                    else:
                        setattr(self._config, key, value)
            
            return self.save_config()
        except Exception as e:
            print(f"Error updating config: {e}")
            return False
    
    def reset_to_defaults(self) -> bool:
        """Reset configuration to default values."""
        self._config = self._create_default_config()
        return self.save_config()
    
    def validate_config(self) -> Dict[str, Any]:
        """Validate current configuration and return validation results."""
        if self._config is None:
            return {"valid": False, "errors": ["No configuration loaded"]}
        
        errors = []
        warnings = []
        
        # Validate Ollama configuration
        if not self._config.ollama.server_url.startswith(('http://', 'https://')):
            errors.append("Invalid Ollama server URL format")
        
        if self._config.ollama.timeout <= 0:
            errors.append("Ollama timeout must be positive")
        
        if self._config.ollama.max_retries < 0:
            errors.append("Max retries cannot be negative")
        
        # Validate Excel settings
        if self._config.excel_settings.max_rows_per_chunk <= 0:
            errors.append("Max rows per chunk must be positive")
        
        # Validate agent settings
        if not 0 <= self._config.agent_settings.analysis_confidence_threshold <= 1:
            errors.append("Analysis confidence threshold must be between 0 and 1")
        
        if not 0 <= self._config.agent_settings.pattern_detection_sensitivity <= 1:
            errors.append("Pattern detection sensitivity must be between 0 and 1")
        
        if self._config.agent_settings.max_concurrent_agents <= 0:
            errors.append("Max concurrent agents must be positive")
        
        # Add warnings for potentially problematic settings
        if self._config.ollama.timeout > 600:
            warnings.append("Ollama timeout is very high (>10 minutes)")
        
        if self._config.excel_settings.max_rows_per_chunk > 50000:
            warnings.append("Large chunk size may cause memory issues")
        
        return {
            "valid": len(errors) == 0,
            "errors": errors,
            "warnings": warnings
        }
    
    def _create_default_config(self) -> PluginConfig:
        """Create default configuration."""
        return PluginConfig(
            ollama=OllamaConfig(),
            excel_settings=ExcelSettings(),
            agent_settings=AgentSettings(),
            ui_preferences=UIPreferences()
        )
    
    def get_config_path(self) -> str:
        """Get path to configuration file."""
        return str(self.config_file)
    
    def backup_config(self) -> bool:
        """Create backup of current configuration."""
        if not self.config_file.exists():
            return False
        
        try:
            backup_file = self.config_dir / f'config_backup_{int(os.path.getmtime(self.config_file))}.json'
            import shutil
            shutil.copy2(self.config_file, backup_file)
            return True
        except Exception as e:
            print(f"Error creating config backup: {e}")
            return False


# Global configuration manager instance
config_manager = ConfigManager()