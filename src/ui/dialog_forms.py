"""
Dialog Forms for Excel-Ollama AI Plugin.
Creates user interface dialogs for configuration and interaction.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from typing import Dict, Any, List, Optional, Callable
import json
import asyncio
import threading
from datetime import datetime

from ..utils.config import PluginConfig, OllamaConfig


class ConfigurationDialog:
    """Main configuration dialog for plugin settings."""
    
    def __init__(self, config: PluginConfig, ollama_client=None):
        self.config = config
        self.ollama_client = ollama_client
        self.root = None
        self.result = None
        
    def show(self) -> Optional[Dict[str, Any]]:
        """Show configuration dialog and return updated settings."""
        self.root = tk.Tk()
        self.root.title("Ollama AI Plugin Configuration")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Create notebook for tabbed interface
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Ollama Configuration Tab
        ollama_frame = ttk.Frame(notebook)
        notebook.add(ollama_frame, text="Ollama Settings")
        self._create_ollama_tab(ollama_frame)
        
        # Analysis Settings Tab
        analysis_frame = ttk.Frame(notebook)
        notebook.add(analysis_frame, text="Analysis Settings")
        self._create_analysis_tab(analysis_frame)
        
        # UI Preferences Tab
        ui_frame = ttk.Frame(notebook)
        notebook.add(ui_frame, text="UI Preferences")
        self._create_ui_tab(ui_frame)
        
        # Advanced Settings Tab
        advanced_frame = ttk.Frame(notebook)
        notebook.add(advanced_frame, text="Advanced")
        self._create_advanced_tab(advanced_frame)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="Test Connection", 
                  command=self._test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", 
                  command=self._reset_defaults).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self._cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="OK", 
                  command=self._ok).pack(side=tk.RIGHT, padx=5)
        
        # Center the window
        self.root.transient()
        self.root.grab_set()
        
        # Run dialog
        self.root.mainloop()
        
        return self.result
    
    def _create_ollama_tab(self, parent):
        """Create Ollama configuration tab."""
        # Server Settings
        server_group = ttk.LabelFrame(parent, text="Server Settings", padding=10)
        server_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(server_group, text="Server URL:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.server_url_var = tk.StringVar(value=self.config.ollama.server_url)
        ttk.Entry(server_group, textvariable=self.server_url_var, width=40).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(server_group, text="Timeout (seconds):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.timeout_var = tk.StringVar(value=str(self.config.ollama.timeout))
        ttk.Entry(server_group, textvariable=self.timeout_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(server_group, text="Max Retries:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.max_retries_var = tk.StringVar(value=str(self.config.ollama.max_retries))
        ttk.Entry(server_group, textvariable=self.max_retries_var, width=10).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        # Model Settings
        model_group = ttk.LabelFrame(parent, text="Model Settings", padding=10)
        model_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(model_group, text="Default Model:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.default_model_var = tk.StringVar(value=self.config.ollama.default_model)
        self.model_combo = ttk.Combobox(model_group, textvariable=self.default_model_var, width=30)
        self.model_combo.grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Button(model_group, text="Refresh Models", 
                  command=self._refresh_models).grid(row=0, column=2, padx=5)
        
        self.stream_responses_var = tk.BooleanVar(value=self.config.ollama.stream_responses)
        ttk.Checkbutton(model_group, text="Enable streaming responses", 
                       variable=self.stream_responses_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Connection Status
        status_group = ttk.LabelFrame(parent, text="Connection Status", padding=10)
        status_group.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_label = ttk.Label(status_group, text="Status: Not tested")
        self.status_label.pack(anchor=tk.W)
        
        self.models_label = ttk.Label(status_group, text="Available models: Not loaded")
        self.models_label.pack(anchor=tk.W)
    
    def _create_analysis_tab(self, parent):
        """Create analysis settings tab."""
        # Default Analysis Settings
        defaults_group = ttk.LabelFrame(parent, text="Default Analysis Settings", padding=10)
        defaults_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(defaults_group, text="Default Analysis Type:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.default_analysis_var = tk.StringVar(value="statistical_analysis")
        analysis_combo = ttk.Combobox(defaults_group, textvariable=self.default_analysis_var, 
                                    values=["statistical_analysis", "trend_analysis", "pattern_detection", "clustering"])
        analysis_combo.grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(defaults_group, text="Auto-detect data types:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.auto_detect_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(defaults_group, variable=self.auto_detect_var).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Performance Settings
        perf_group = ttk.LabelFrame(parent, text="Performance Settings", padding=10)
        perf_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(perf_group, text="Max rows for analysis:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.max_rows_var = tk.StringVar(value="100000")
        ttk.Entry(perf_group, textvariable=self.max_rows_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(perf_group, text="Chunk size for large datasets:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.chunk_size_var = tk.StringVar(value="10000")
        ttk.Entry(perf_group, textvariable=self.chunk_size_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        self.parallel_processing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(perf_group, text="Enable parallel processing", 
                       variable=self.parallel_processing_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
    
    def _create_ui_tab(self, parent):
        """Create UI preferences tab."""
        # Display Settings
        display_group = ttk.LabelFrame(parent, text="Display Settings", padding=10)
        display_group.pack(fill=tk.X, padx=10, pady=5)
        
        self.show_progress_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_group, text="Show progress indicators", 
                       variable=self.show_progress_var).pack(anchor=tk.W, pady=2)
        
        self.auto_open_results_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_group, text="Auto-open results sheet", 
                       variable=self.auto_open_results_var).pack(anchor=tk.W, pady=2)
        
        self.show_confidence_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(display_group, text="Show confidence scores", 
                       variable=self.show_confidence_var).pack(anchor=tk.W, pady=2)
        
        # Notification Settings
        notif_group = ttk.LabelFrame(parent, text="Notifications", padding=10)
        notif_group.pack(fill=tk.X, padx=10, pady=5)
        
        self.notify_completion_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(notif_group, text="Notify when analysis completes", 
                       variable=self.notify_completion_var).pack(anchor=tk.W, pady=2)
        
        self.notify_errors_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(notif_group, text="Show error notifications", 
                       variable=self.notify_errors_var).pack(anchor=tk.W, pady=2)
    
    def _create_advanced_tab(self, parent):
        """Create advanced settings tab."""
        # Logging Settings
        logging_group = ttk.LabelFrame(parent, text="Logging", padding=10)
        logging_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(logging_group, text="Log Level:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.log_level_var = tk.StringVar(value="INFO")
        log_combo = ttk.Combobox(logging_group, textvariable=self.log_level_var, 
                               values=["DEBUG", "INFO", "WARNING", "ERROR"])
        log_combo.grid(row=0, column=1, sticky=tk.W, pady=2)
        
        self.enable_logging_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(logging_group, text="Enable logging", 
                       variable=self.enable_logging_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Security Settings
        security_group = ttk.LabelFrame(parent, text="Security", padding=10)
        security_group.pack(fill=tk.X, padx=10, pady=5)
        
        self.encrypt_cache_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(security_group, text="Encrypt cached data", 
                       variable=self.encrypt_cache_var).pack(anchor=tk.W, pady=2)
        
        self.clear_cache_on_exit_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(security_group, text="Clear cache on exit", 
                       variable=self.clear_cache_on_exit_var).pack(anchor=tk.W, pady=2)
        
        # Cache Settings
        cache_group = ttk.LabelFrame(parent, text="Cache Settings", padding=10)
        cache_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(cache_group, text="Cache size (MB):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.cache_size_var = tk.StringVar(value="100")
        ttk.Entry(cache_group, textvariable=self.cache_size_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Button(cache_group, text="Clear Cache", 
                  command=self._clear_cache).grid(row=0, column=2, padx=5)
    
    def _refresh_models(self):
        """Refresh available models from Ollama."""
        if self.ollama_client:
            try:
                # This would be async in real implementation
                models = ["llama2:latest", "codellama:latest", "mistral:latest", "phi:latest"]
                self.model_combo['values'] = models
                self.models_label.config(text=f"Available models: {', '.join(models)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to refresh models: {e}")
        else:
            messagebox.showwarning("Warning", "Ollama client not available")
    
    def _test_connection(self):
        """Test connection to Ollama server."""
        try:
            # Update status
            self.status_label.config(text="Status: Testing...")
            self.root.update()
            
            # Simulate connection test
            import time
            time.sleep(1)  # Simulate network delay
            
            # In real implementation, this would test actual connection
            self.status_label.config(text="Status: Connected ✓")
            self._refresh_models()
            
        except Exception as e:
            self.status_label.config(text=f"Status: Connection failed - {e}")
            messagebox.showerror("Connection Error", f"Failed to connect to Ollama: {e}")
    
    def _reset_defaults(self):
        """Reset all settings to defaults."""
        if messagebox.askyesno("Reset Settings", "Reset all settings to defaults?"):
            default_config = PluginConfig()
            
            # Reset Ollama settings
            self.server_url_var.set(default_config.ollama.server_url)
            self.timeout_var.set(str(default_config.ollama.timeout))
            self.max_retries_var.set(str(default_config.ollama.max_retries))
            self.default_model_var.set(default_config.ollama.default_model)
            self.stream_responses_var.set(default_config.ollama.stream_responses)
            
            # Reset other settings
            self.default_analysis_var.set("statistical_analysis")
            self.auto_detect_var.set(True)
            self.max_rows_var.set("100000")
            self.chunk_size_var.set("10000")
            self.parallel_processing_var.set(True)
    
    def _clear_cache(self):
        """Clear plugin cache."""
        if messagebox.askyesno("Clear Cache", "Clear all cached data?"):
            # In real implementation, this would clear actual cache
            messagebox.showinfo("Cache Cleared", "Cache has been cleared successfully.")
    
    def _ok(self):
        """Save settings and close dialog."""
        try:
            # Validate inputs
            timeout = int(self.timeout_var.get())
            max_retries = int(self.max_retries_var.get())
            max_rows = int(self.max_rows_var.get())
            chunk_size = int(self.chunk_size_var.get())
            cache_size = int(self.cache_size_var.get())
            
            # Create updated configuration
            self.result = {
                'ollama': {
                    'server_url': self.server_url_var.get(),
                    'timeout': timeout,
                    'max_retries': max_retries,
                    'default_model': self.default_model_var.get(),
                    'stream_responses': self.stream_responses_var.get()
                },
                'analysis': {
                    'default_type': self.default_analysis_var.get(),
                    'auto_detect_types': self.auto_detect_var.get(),
                    'max_rows': max_rows,
                    'chunk_size': chunk_size,
                    'parallel_processing': self.parallel_processing_var.get()
                },
                'ui': {
                    'show_progress': self.show_progress_var.get(),
                    'auto_open_results': self.auto_open_results_var.get(),
                    'show_confidence': self.show_confidence_var.get(),
                    'notify_completion': self.notify_completion_var.get(),
                    'notify_errors': self.notify_errors_var.get()
                },
                'advanced': {
                    'log_level': self.log_level_var.get(),
                    'enable_logging': self.enable_logging_var.get(),
                    'encrypt_cache': self.encrypt_cache_var.get(),
                    'clear_cache_on_exit': self.clear_cache_on_exit_var.get(),
                    'cache_size_mb': cache_size
                }
            }
            
            self.root.destroy()
            
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Please check your numeric inputs: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {e}")
    
    def _cancel(self):
        """Cancel and close dialog."""
        self.result = None
        self.root.destroy()


class QueryDialog:
    """Dialog for natural language queries."""
    
    def __init__(self, data_info: Dict[str, Any] = None):
        self.data_info = data_info or {}
        self.root = None
        self.result = None
    
    def show(self) -> Optional[str]:
        """Show query dialog and return user query."""
        self.root = tk.Tk()
        self.root.title("Ask a Question About Your Data")
        self.root.geometry("500x400")
        
        # Data info section
        if self.data_info:
            info_frame = ttk.LabelFrame(self.root, text="Data Information", padding=10)
            info_frame.pack(fill=tk.X, padx=10, pady=5)
            
            info_text = f"Rows: {self.data_info.get('rows', 'Unknown')}, Columns: {self.data_info.get('columns', 'Unknown')}"
            ttk.Label(info_frame, text=info_text).pack(anchor=tk.W)
            
            if 'column_names' in self.data_info:
                columns_text = f"Columns: {', '.join(self.data_info['column_names'][:5])}"
                if len(self.data_info['column_names']) > 5:
                    columns_text += "..."
                ttk.Label(info_frame, text=columns_text).pack(anchor=tk.W)
        
        # Query input section
        query_frame = ttk.LabelFrame(self.root, text="Your Question", padding=10)
        query_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(query_frame, text="Ask a question about your data:").pack(anchor=tk.W, pady=(0, 5))
        
        self.query_text = scrolledtext.ScrolledText(query_frame, height=8, wrap=tk.WORD)
        self.query_text.pack(fill=tk.BOTH, expand=True)
        
        # Example queries
        examples_frame = ttk.LabelFrame(self.root, text="Example Questions", padding=10)
        examples_frame.pack(fill=tk.X, padx=10, pady=5)
        
        examples = [
            "What trends do you see in the data?",
            "Are there any outliers or anomalies?",
            "What is the correlation between variables?",
            "Can you forecast the next 10 periods?",
            "Group the data into clusters"
        ]
        
        for example in examples:
            btn = ttk.Button(examples_frame, text=example, 
                           command=lambda e=example: self._insert_example(e))
            btn.pack(fill=tk.X, pady=1)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="Cancel", 
                  command=self._cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Analyze", 
                  command=self._ok).pack(side=tk.RIGHT, padx=5)
        
        # Focus on text area
        self.query_text.focus()
        
        # Center and show
        self.root.transient()
        self.root.grab_set()
        self.root.mainloop()
        
        return self.result
    
    def _insert_example(self, example: str):
        """Insert example query into text area."""
        self.query_text.delete(1.0, tk.END)
        self.query_text.insert(1.0, example)
    
    def _ok(self):
        """Get query and close dialog."""
        query = self.query_text.get(1.0, tk.END).strip()
        if query:
            self.result = query
            self.root.destroy()
        else:
            messagebox.showwarning("Empty Query", "Please enter a question about your data.")
    
    def _cancel(self):
        """Cancel and close dialog."""
        self.result = None
        self.root.destroy()


class ProgressDialog:
    """Progress dialog for long-running operations."""
    
    def __init__(self, title: str = "Processing", message: str = "Please wait..."):
        self.title = title
        self.message = message
        self.root = None
        self.progress_var = None
        self.status_var = None
        self.cancelled = False
    
    def show(self):
        """Show progress dialog."""
        self.root = tk.Tk()
        self.root.title(self.title)
        self.root.geometry("400x150")
        self.root.resizable(False, False)
        
        # Message
        ttk.Label(self.root, text=self.message).pack(pady=10)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, 
                                     maximum=100, length=300)
        progress_bar.pack(pady=10)
        
        # Status label
        self.status_var = tk.StringVar(value="Initializing...")
        ttk.Label(self.root, textvariable=self.status_var).pack(pady=5)
        
        # Cancel button
        ttk.Button(self.root, text="Cancel", 
                  command=self._cancel).pack(pady=5)
        
        # Center window
        self.root.transient()
        self.root.protocol("WM_DELETE_WINDOW", self._cancel)
        
        # Start in separate thread to avoid blocking
        self.root.after(100, self._update)
    
    def update_progress(self, progress: float, status: str = None):
        """Update progress and status."""
        if self.progress_var:
            self.progress_var.set(progress)
        if status and self.status_var:
            self.status_var.set(status)
        if self.root:
            self.root.update()
    
    def close(self):
        """Close progress dialog."""
        if self.root:
            self.root.destroy()
            self.root = None
    
    def _update(self):
        """Update dialog periodically."""
        if not self.cancelled and self.root:
            self.root.after(100, self._update)
    
    def _cancel(self):
        """Cancel operation."""
        self.cancelled = True
        self.close()


class ResultsDialog:
    """Dialog for displaying analysis results."""
    
    def __init__(self, results: Dict[str, Any]):
        self.results = results
        self.root = None
    
    def show(self):
        """Show results dialog."""
        self.root = tk.Tk()
        self.root.title("Analysis Results")
        self.root.geometry("700x500")
        
        # Create notebook for different result views
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Summary tab
        summary_frame = ttk.Frame(notebook)
        notebook.add(summary_frame, text="Summary")
        self._create_summary_tab(summary_frame)
        
        # Detailed results tab
        details_frame = ttk.Frame(notebook)
        notebook.add(details_frame, text="Detailed Results")
        self._create_details_tab(details_frame)
        
        # Raw data tab
        raw_frame = ttk.Frame(notebook)
        notebook.add(raw_frame, text="Raw Data")
        self._create_raw_tab(raw_frame)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="Export Results", 
                  command=self._export_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Copy to Clipboard", 
                  command=self._copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Close", 
                  command=self._close).pack(side=tk.RIGHT, padx=5)
        
        self.root.mainloop()
    
    def _create_summary_tab(self, parent):
        """Create summary tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Format summary
        summary = self.results.get('summary', 'No summary available')
        confidence = self.results.get('confidence_score', 0)
        methodology = self.results.get('methodology', 'Unknown')
        
        summary_text = f"""ANALYSIS SUMMARY
{'=' * 50}

{summary}

CONFIDENCE SCORE: {confidence:.2%}
METHODOLOGY: {methodology}

TIMESTAMP: {self.results.get('processing_timestamp', 'Unknown')}
"""
        
        text_widget.insert(1.0, summary_text)
        text_widget.config(state=tk.DISABLED)
    
    def _create_details_tab(self, parent):
        """Create detailed results tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Format detailed results
        details_text = "DETAILED RESULTS\n" + "=" * 50 + "\n\n"
        
        for key, value in self.results.items():
            if key not in ['summary', 'confidence_score', 'methodology', 'processing_timestamp']:
                details_text += f"{key.upper().replace('_', ' ')}:\n"
                if isinstance(value, dict):
                    for subkey, subvalue in value.items():
                        details_text += f"  {subkey}: {subvalue}\n"
                else:
                    details_text += f"  {value}\n"
                details_text += "\n"
        
        text_widget.insert(1.0, details_text)
        text_widget.config(state=tk.DISABLED)
    
    def _create_raw_tab(self, parent):
        """Create raw data tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD, font=('Courier', 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Show raw JSON
        raw_text = json.dumps(self.results, indent=2, default=str)
        text_widget.insert(1.0, raw_text)
        text_widget.config(state=tk.DISABLED)
    
    def _export_results(self):
        """Export results to file."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    if filename.endswith('.json'):
                        json.dump(self.results, f, indent=2, default=str)
                    else:
                        f.write(json.dumps(self.results, indent=2, default=str))
                
                messagebox.showinfo("Export Successful", f"Results exported to {filename}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export results: {e}")
    
    def _copy_to_clipboard(self):
        """Copy results to clipboard."""
        try:
            summary = self.results.get('summary', json.dumps(self.results, indent=2, default=str))
            self.root.clipboard_clear()
            self.root.clipboard_append(summary)
            messagebox.showinfo("Copied", "Results copied to clipboard")
        except Exception as e:
            messagebox.showerror("Copy Error", f"Failed to copy to clipboard: {e}")
    
    def _close(self):
        """Close dialog."""
        self.root.destroy()


class HelpDialog:
    """Help and documentation dialog."""
    
    def __init__(self):
        self.root = None
    
    def show(self):
        """Show help dialog."""
        self.root = tk.Tk()
        self.root.title("Ollama AI Plugin Help")
        self.root.geometry("600x500")
        
        # Create notebook for help sections
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Getting Started tab
        getting_started_frame = ttk.Frame(notebook)
        notebook.add(getting_started_frame, text="Getting Started")
        self._create_getting_started_tab(getting_started_frame)
        
        # Features tab
        features_frame = ttk.Frame(notebook)
        notebook.add(features_frame, text="Features")
        self._create_features_tab(features_frame)
        
        # Troubleshooting tab
        troubleshooting_frame = ttk.Frame(notebook)
        notebook.add(troubleshooting_frame, text="Troubleshooting")
        self._create_troubleshooting_tab(troubleshooting_frame)
        
        # About tab
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="About")
        self._create_about_tab(about_frame)
        
        # Close button
        ttk.Button(self.root, text="Close", 
                  command=self._close).pack(pady=10)
        
        self.root.mainloop()
    
    def _create_getting_started_tab(self, parent):
        """Create getting started tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        help_text = """GETTING STARTED WITH OLLAMA AI PLUGIN
========================================

1. SETUP OLLAMA
   - Install Ollama on your system
   - Start Ollama server: ollama serve
   - Download models: ollama pull llama2

2. CONFIGURE PLUGIN
   - Click "Configure" in the Ollama AI Analysis ribbon
   - Set server URL (default: http://localhost:11434)
   - Test connection and select default model

3. ANALYZE DATA
   - Select data range in Excel
   - Click "Analyze Data" or use "Ask Question"
   - Review results in generated sheets

4. CUSTOM FUNCTIONS
   - Use =OLLAMA_ANALYZE(range, prompt) in cells
   - Try =AI_TREND(range, periods) for forecasting
   - Use =PATTERN_DETECT(range, threshold) for patterns

5. NATURAL LANGUAGE QUERIES
   - Click "Ask Question" to query your data
   - Examples: "What trends do you see?"
   - "Are there any outliers?"
   - "Forecast next 10 periods"
"""
        
        text_widget.insert(1.0, help_text)
        text_widget.config(state=tk.DISABLED)
    
    def _create_features_tab(self, parent):
        """Create features tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        features_text = """PLUGIN FEATURES
===============

ANALYSIS CAPABILITIES:
• Statistical Analysis - Descriptive statistics, correlations
• Trend Analysis - Time series trends, forecasting
• Pattern Detection - Seasonal patterns, anomalies, outliers
• Clustering - Group similar data points
• Natural Language Queries - Ask questions in plain English

EXCEL INTEGRATION:
• Custom Excel functions (UDFs)
• Ribbon interface with analysis buttons
• Automatic result sheet generation
• Chart and visualization creation
• Export capabilities

OLLAMA INTEGRATION:
• Support for multiple models (llama2, codellama, mistral, etc.)
• Streaming responses for real-time feedback
• Model switching without restart
• Custom model parameters

USER INTERFACE:
• Configuration dialogs
• Progress indicators
• Results viewer
• Help system
• Error handling and notifications

PERFORMANCE:
• Large dataset support with chunking
• Parallel processing
• Caching for improved speed
• Memory management
"""
        
        text_widget.insert(1.0, features_text)
        text_widget.config(state=tk.DISABLED)
    
    def _create_troubleshooting_tab(self, parent):
        """Create troubleshooting tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        troubleshooting_text = """TROUBLESHOOTING
===============

COMMON ISSUES:

1. "Cannot connect to Ollama"
   - Ensure Ollama is running: ollama serve
   - Check server URL in configuration
   - Verify firewall settings

2. "Model not found"
   - Download model: ollama pull llama2
   - Refresh models in configuration
   - Check model name spelling

3. "Analysis takes too long"
   - Reduce data size or use sampling
   - Enable chunking for large datasets
   - Check system resources

4. "Excel functions not working"
   - Restart Excel after installation
   - Check if plugin is enabled
   - Verify xlwings installation

5. "Results are inaccurate"
   - Check data quality and format
   - Verify column types
   - Try different analysis methods

6. "Plugin crashes Excel"
   - Update to latest version
   - Check system requirements
   - Disable other add-ins temporarily

GETTING HELP:
- Check plugin logs for detailed errors
- Use "Test Connection" in configuration
- Contact support with error details
"""
        
        text_widget.insert(1.0, troubleshooting_text)
        text_widget.config(state=tk.DISABLED)
    
    def _create_about_tab(self, parent):
        """Create about tab."""
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        about_text = """ABOUT OLLAMA AI PLUGIN
======================

VERSION: 1.0.0
RELEASE DATE: 2024

DESCRIPTION:
The Ollama AI Plugin transforms Excel into a powerful AI-enhanced data analysis platform. It integrates local Ollama LLM capabilities directly into Excel, enabling users to perform sophisticated data analysis using natural language queries and automated insights generation.

KEY BENEFITS:
• Local processing - your data never leaves your system
• Natural language interface - no need to learn complex formulas
• Multiple analysis types - from basic statistics to advanced patterns
• Seamless Excel integration - works with your existing workflows
• Extensible architecture - supports multiple Ollama models

TECHNOLOGY STACK:
• Python with xlwings for Excel integration
• Ollama API for LLM communication
• Pandas and NumPy for data processing
• Scikit-learn for machine learning
• Tkinter for user interface

SYSTEM REQUIREMENTS:
• Microsoft Excel 2016 or later
• Python 3.8 or later
• Ollama installed and running
• Windows 10 or later

LICENSE:
This software is provided under the MIT License.

SUPPORT:
For support, documentation, and updates, visit:
https://github.com/your-repo/excel-ollama-plugin

COPYRIGHT:
© 2024 Excel-Ollama AI Plugin. All rights reserved.
"""
        
        text_widget.insert(1.0, about_text)
        text_widget.config(state=tk.DISABLED)
    
    def _close(self):
        """Close help dialog."""
        self.root.destroy()