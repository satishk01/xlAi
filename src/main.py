"""
Main entry point for Excel-Ollama AI Plugin.
Handles Excel integration and ribbon callbacks.
"""

import sys
import os
import asyncio
import logging
from typing import Dict, Any, Optional
import xlwings as xw

# Add src directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.ollama_client import OllamaClient
from core.agent_controller import AgentController
from core.excel_interface import ExcelInterface
from core.query_processor import QueryProcessor, QueryResponseFormatter
from core.data_processor import DataProcessor
from agents.analysis_agent import AnalysisAgent
from agents.pattern_agent import PatternAgent
from agents.reporting_agent import ReportingAgent
from ui.dialog_forms import ConfigurationDialog, QueryDialog, ProgressDialog, ResultsDialog, HelpDialog
from utils.config import PluginConfig
from utils.logger import setup_logging


class ExcelOllamaPlugin:
    """Main plugin class that coordinates all components."""
    
    def __init__(self):
        self.config = PluginConfig()
        self.logger = setup_logging(self.config.advanced.get('log_level', 'INFO'))
        
        # Core components
        self.ollama_client = None
        self.agent_controller = None
        self.excel_interface = None
        self.query_processor = None
        self.query_formatter = None
        self.data_processor = None
        
        # UI components
        self.progress_dialog = None
        
        # Initialize plugin
        self._initialize_plugin()
    
    def _initialize_plugin(self):
        """Initialize all plugin components."""
        try:
            self.logger.info("Initializing Excel-Ollama AI Plugin...")
            
            # Initialize core components
            self.ollama_client = OllamaClient(self.config.ollama.server_url)
            self.data_processor = DataProcessor()
            self.excel_interface = ExcelInterface(self.config)
            self.query_processor = QueryProcessor(self.ollama_client)
            self.query_formatter = QueryResponseFormatter(self.ollama_client)
            
            # Initialize agents
            analysis_agent = AnalysisAgent(self.ollama_client)
            pattern_agent = PatternAgent(self.ollama_client)
            reporting_agent = ReportingAgent(self.ollama_client)
            
            # Initialize agent controller
            self.agent_controller = AgentController(self.ollama_client)
            self.agent_controller.register_agent('analysis', analysis_agent)
            self.agent_controller.register_agent('pattern', pattern_agent)
            self.agent_controller.register_agent('reporting', reporting_agent)
            
            # Register Excel functions
            self.excel_interface.register_custom_functions()
            
            self.logger.info("Plugin initialized successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to initialize plugin: {e}")
            raise
    
    # Ribbon callback functions
    def OnRibbonLoad(self, ribbon):
        """Called when ribbon is loaded."""
        self.logger.info("Ribbon loaded")
        self.ribbon = ribbon
        return True
    
    def OnAnalyzeData(self, control):
        """Handle Analyze Data button click."""
        try:
            self.logger.info("Analyze Data button clicked")
            
            # Get selected data
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range to analyze.")
                return
            
            # Show progress
            self.progress_dialog = ProgressDialog("Analyzing Data", "Performing statistical analysis...")
            self.progress_dialog.show()
            
            # Run analysis in background
            asyncio.create_task(self._run_analysis(data, 'statistical_analysis'))
            
        except Exception as e:
            self.logger.error(f"Error in OnAnalyzeData: {e}")
            self._show_error(f"Analysis failed: {e}")
    
    def OnTrendAnalysis(self, control):
        """Handle Trend Analysis button click."""
        try:
            self.logger.info("Trend Analysis button clicked")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range for trend analysis.")
                return
            
            # Check for time series data
            datetime_cols = data.select_dtypes(include=['datetime64']).columns
            if len(datetime_cols) == 0:
                self._show_error("Trend analysis requires a datetime column. Please ensure your data includes dates/times.")
                return
            
            self.progress_dialog = ProgressDialog("Trend Analysis", "Analyzing trends and patterns...")
            self.progress_dialog.show()
            
            asyncio.create_task(self._run_analysis(data, 'trend_analysis'))
            
        except Exception as e:
            self.logger.error(f"Error in OnTrendAnalysis: {e}")
            self._show_error(f"Trend analysis failed: {e}")
    
    def OnPatternDetection(self, control):
        """Handle Pattern Detection button click."""
        try:
            self.logger.info("Pattern Detection button clicked")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range for pattern detection.")
                return
            
            self.progress_dialog = ProgressDialog("Pattern Detection", "Detecting patterns and anomalies...")
            self.progress_dialog.show()
            
            asyncio.create_task(self._run_analysis(data, 'pattern_detection'))
            
        except Exception as e:
            self.logger.error(f"Error in OnPatternDetection: {e}")
            self._show_error(f"Pattern detection failed: {e}")
    
    def OnQueryData(self, control):
        """Handle Ask Question button click."""
        try:
            self.logger.info("Query Data button clicked")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range to query.")
                return
            
            # Show query dialog
            data_info = {
                'rows': len(data),
                'columns': len(data.columns),
                'column_names': list(data.columns)
            }
            
            query_dialog = QueryDialog(data_info)
            query = query_dialog.show()
            
            if query:
                self.progress_dialog = ProgressDialog("Processing Query", "Understanding your question...")
                self.progress_dialog.show()
                
                asyncio.create_task(self._process_query(query, data))
            
        except Exception as e:
            self.logger.error(f"Error in OnQueryData: {e}")
            self._show_error(f"Query processing failed: {e}")
    
    def OnQuickQuery(self, control, text):
        """Handle quick query input."""
        try:
            if not text.strip():
                return
            
            self.logger.info(f"Quick query: {text}")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range first.")
                return
            
            self.progress_dialog = ProgressDialog("Processing Query", "Analyzing your question...")
            self.progress_dialog.show()
            
            asyncio.create_task(self._process_query(text, data))
            
        except Exception as e:
            self.logger.error(f"Error in OnQuickQuery: {e}")
            self._show_error(f"Quick query failed: {e}")
    
    def OnGenerateReport(self, control):
        """Handle Generate Report button click."""
        try:
            self.logger.info("Generate Report button clicked")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range for report generation.")
                return
            
            self.progress_dialog = ProgressDialog("Generating Report", "Creating comprehensive analysis report...")
            self.progress_dialog.show()
            
            asyncio.create_task(self._generate_report(data))
            
        except Exception as e:
            self.logger.error(f"Error in OnGenerateReport: {e}")
            self._show_error(f"Report generation failed: {e}")
    
    def OnCreateDashboard(self, control):
        """Handle Create Dashboard button click."""
        try:
            self.logger.info("Create Dashboard button clicked")
            
            data = self.excel_interface.get_selected_range()
            if data is None or data.empty:
                self._show_error("Please select a data range for dashboard creation.")
                return
            
            self.progress_dialog = ProgressDialog("Creating Dashboard", "Building executive dashboard...")
            self.progress_dialog.show()
            
            asyncio.create_task(self._create_dashboard(data))
            
        except Exception as e:
            self.logger.error(f"Error in OnCreateDashboard: {e}")
            self._show_error(f"Dashboard creation failed: {e}")
    
    def OnExportResults(self, control):
        """Handle Export Results button click."""
        try:
            self.logger.info("Export Results button clicked")
            
            # Get data from results sheet
            try:
                results_data = self.excel_interface.get_worksheet_data("AI_Analysis_Results")
                if results_data is None or results_data.empty:
                    self._show_error("No analysis results found to export.")
                    return
                
                # Export to file
                success = self.excel_interface.export_results(results_data, 'xlsx')
                if success:
                    self._show_info("Results exported successfully.")
                else:
                    self._show_error("Failed to export results.")
                    
            except Exception as e:
                self._show_error(f"No results sheet found. Please run an analysis first.")
            
        except Exception as e:
            self.logger.error(f"Error in OnExportResults: {e}")
            self._show_error(f"Export failed: {e}")
    
    def OnModelSelection(self, control, selectedId, selectedIndex):
        """Handle model selection dropdown."""
        try:
            self.logger.info(f"Model selected: {selectedId}")
            
            # Update configuration
            self.config.ollama.default_model = selectedId
            
            # Update Ollama client
            if self.ollama_client:
                asyncio.create_task(self.ollama_client.load_model(selectedId))
            
            self.excel_interface.update_ribbon_status(f"Model: {selectedId}")
            
        except Exception as e:
            self.logger.error(f"Error in OnModelSelection: {e}")
            self._show_error(f"Model selection failed: {e}")
    
    def OnConfigure(self, control):
        """Handle Configure button click."""
        try:
            self.logger.info("Configure button clicked")
            
            config_dialog = ConfigurationDialog(self.config, self.ollama_client)
            new_config = config_dialog.show()
            
            if new_config:
                # Update configuration
                self.config.update_from_dict(new_config)
                
                # Reinitialize components with new config
                self._reinitialize_with_config()
                
                self._show_info("Configuration updated successfully.")
            
        except Exception as e:
            self.logger.error(f"Error in OnConfigure: {e}")
            self._show_error(f"Configuration failed: {e}")
    
    def OnHelp(self, control):
        """Handle Help button click."""
        try:
            self.logger.info("Help button clicked")
            
            help_dialog = HelpDialog()
            help_dialog.show()
            
        except Exception as e:
            self.logger.error(f"Error in OnHelp: {e}")
            self._show_error(f"Help display failed: {e}")
    
    def OnRefreshConnection(self, control):
        """Handle Refresh Connection button click."""
        try:
            self.logger.info("Refresh Connection button clicked")
            
            # Test connection
            asyncio.create_task(self._test_connection())
            
        except Exception as e:
            self.logger.error(f"Error in OnRefreshConnection: {e}")
            self._show_error(f"Connection refresh failed: {e}")
    
    # Async analysis methods
    async def _run_analysis(self, data, analysis_type):
        """Run analysis in background."""
        try:
            self.progress_dialog.update_progress(10, "Preparing data...")
            
            # Validate and process data
            processed_data = self.data_processor.validate_data(data)
            if not processed_data.is_valid:
                self._show_error(f"Data validation failed: {processed_data.errors}")
                return
            
            self.progress_dialog.update_progress(30, "Running analysis...")
            
            # Run analysis through agent controller
            result = await self.agent_controller.execute_analysis_pipeline(
                data, analysis_type
            )
            
            self.progress_dialog.update_progress(80, "Formatting results...")
            
            # Write results to Excel
            self.excel_interface.write_results_to_sheet(result, f"AI_{analysis_type}_Results")
            
            self.progress_dialog.update_progress(100, "Complete!")
            self.progress_dialog.close()
            
            # Show results dialog
            results_dialog = ResultsDialog(result)
            results_dialog.show()
            
        except Exception as e:
            self.logger.error(f"Analysis failed: {e}")
            if self.progress_dialog:
                self.progress_dialog.close()
            self._show_error(f"Analysis failed: {e}")
    
    async def _process_query(self, query, data):
        """Process natural language query."""
        try:
            self.progress_dialog.update_progress(20, "Understanding query...")
            
            # Process query
            query_result = await self.query_processor.process_query(query, data)
            
            if 'error' in query_result:
                self._show_error(f"Query processing failed: {query_result['error']}")
                return
            
            self.progress_dialog.update_progress(50, "Executing analysis...")
            
            # Execute analysis based on query
            analysis_spec = query_result['analysis_specification']
            result = await self.agent_controller.execute_analysis_pipeline(
                data, analysis_spec['method'], analysis_spec['parameters']
            )
            
            self.progress_dialog.update_progress(80, "Formatting response...")
            
            # Format response
            formatted_response = await self.query_formatter.format_response(query, result)
            
            # Combine query processing and analysis results
            combined_result = {
                'original_query': query,
                'query_processing': query_result,
                'analysis_result': result,
                'formatted_response': formatted_response
            }
            
            self.progress_dialog.update_progress(100, "Complete!")
            self.progress_dialog.close()
            
            # Write results
            self.excel_interface.write_results_to_sheet(combined_result, "AI_Query_Results")
            
            # Show results
            results_dialog = ResultsDialog(combined_result)
            results_dialog.show()
            
        except Exception as e:
            self.logger.error(f"Query processing failed: {e}")
            if self.progress_dialog:
                self.progress_dialog.close()
            self._show_error(f"Query processing failed: {e}")
    
    async def _generate_report(self, data):
        """Generate comprehensive report."""
        try:
            self.progress_dialog.update_progress(20, "Analyzing data...")
            
            # Run multiple analyses
            analyses = {}
            
            # Statistical analysis
            analyses['statistics'] = await self.agent_controller.execute_analysis_pipeline(
                data, 'statistical_analysis'
            )
            
            self.progress_dialog.update_progress(40, "Detecting patterns...")
            
            # Pattern analysis
            analyses['patterns'] = await self.agent_controller.execute_analysis_pipeline(
                data, 'pattern_detection'
            )
            
            self.progress_dialog.update_progress(60, "Generating insights...")
            
            # Get reporting agent
            reporting_agent = self.agent_controller.get_agent_by_type('reporting')
            
            # Generate comprehensive summary
            report = await reporting_agent.generate_summary(analyses)
            
            self.progress_dialog.update_progress(80, "Creating report...")
            
            # Create formatted report
            formatted_report = await reporting_agent.create_report('executive_summary', analyses)
            
            combined_result = {
                'report': formatted_report,
                'summary': report,
                'detailed_analyses': analyses
            }
            
            self.progress_dialog.update_progress(100, "Complete!")
            self.progress_dialog.close()
            
            # Write results
            self.excel_interface.write_results_to_sheet(combined_result, "AI_Report")
            
            # Show results
            results_dialog = ResultsDialog(combined_result)
            results_dialog.show()
            
        except Exception as e:
            self.logger.error(f"Report generation failed: {e}")
            if self.progress_dialog:
                self.progress_dialog.close()
            self._show_error(f"Report generation failed: {e}")
    
    async def _create_dashboard(self, data):
        """Create executive dashboard."""
        try:
            self.progress_dialog.update_progress(30, "Extracting key metrics...")
            
            # Run analysis to get key metrics
            analysis_result = await self.agent_controller.execute_analysis_pipeline(
                data, 'statistical_analysis'
            )
            
            self.progress_dialog.update_progress(60, "Building dashboard...")
            
            # Get reporting agent
            reporting_agent = self.agent_controller.get_agent_by_type('reporting')
            
            # Extract key metrics
            key_metrics = reporting_agent._extract_key_metrics(analysis_result)
            
            # Build dashboard
            dashboard = await reporting_agent.build_dashboard(key_metrics)
            
            self.progress_dialog.update_progress(90, "Creating visualizations...")
            
            # Create visualizations in Excel
            data_characteristics = self.excel_interface.get_data_characteristics(data)
            viz_recommendations = await reporting_agent.recommend_visualizations(data_characteristics)
            
            combined_result = {
                'dashboard': dashboard,
                'visualizations': viz_recommendations,
                'key_metrics': key_metrics
            }
            
            self.progress_dialog.update_progress(100, "Complete!")
            self.progress_dialog.close()
            
            # Write results
            self.excel_interface.write_results_to_sheet(combined_result, "AI_Dashboard")
            
            # Create charts
            if not data.empty:
                self.excel_interface.create_visualization('line', data, "AI_Charts")
            
            # Show results
            results_dialog = ResultsDialog(combined_result)
            results_dialog.show()
            
        except Exception as e:
            self.logger.error(f"Dashboard creation failed: {e}")
            if self.progress_dialog:
                self.progress_dialog.close()
            self._show_error(f"Dashboard creation failed: {e}")
    
    async def _test_connection(self):
        """Test connection to Ollama server."""
        try:
            self.excel_interface.update_ribbon_status("Testing connection...")
            
            # Test connection
            models = await self.ollama_client.list_models()
            
            if models:
                self.excel_interface.update_ribbon_status(f"Connected - {len(models)} models available")
                self._show_info(f"Connection successful! Found {len(models)} models.")
            else:
                self.excel_interface.update_ribbon_status("Connected - No models")
                self._show_info("Connected but no models found. Please download models using 'ollama pull'.")
                
        except Exception as e:
            self.logger.error(f"Connection test failed: {e}")
            self.excel_interface.update_ribbon_status("Connection failed")
            self._show_error(f"Connection failed: {e}")
    
    def _reinitialize_with_config(self):
        """Reinitialize components with updated configuration."""
        try:
            # Update Ollama client
            self.ollama_client.base_url = self.config.ollama.server_url
            self.ollama_client.timeout = self.config.ollama.timeout
            
            # Update logging level
            logging.getLogger().setLevel(self.config.advanced.get('log_level', 'INFO'))
            
            self.logger.info("Components reinitialized with new configuration")
            
        except Exception as e:
            self.logger.error(f"Failed to reinitialize components: {e}")
    
    def _show_error(self, message):
        """Show error message to user."""
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", message)
        except:
            # Fallback to Excel status bar
            if self.excel_interface:
                self.excel_interface.update_ribbon_status(f"Error: {message}")
    
    def _show_info(self, message):
        """Show info message to user."""
        try:
            import tkinter.messagebox as messagebox
            messagebox.showinfo("Information", message)
        except:
            # Fallback to Excel status bar
            if self.excel_interface:
                self.excel_interface.update_ribbon_status(message)
    
    def cleanup(self):
        """Clean up plugin resources."""
        try:
            if self.excel_interface:
                self.excel_interface.cleanup()
            
            if self.ollama_client:
                asyncio.create_task(self.ollama_client.close())
            
            self.logger.info("Plugin cleanup completed")
            
        except Exception as e:
            self.logger.error(f"Error during cleanup: {e}")


# Global plugin instance
plugin_instance = None


def initialize_plugin():
    """Initialize the plugin."""
    global plugin_instance
    try:
        plugin_instance = ExcelOllamaPlugin()
        return plugin_instance
    except Exception as e:
        logging.error(f"Failed to initialize plugin: {e}")
        raise


# Excel callback functions (these are called by xlwings)
def OnRibbonLoad(ribbon):
    """Ribbon load callback."""
    if plugin_instance:
        return plugin_instance.OnRibbonLoad(ribbon)
    return False


def OnAnalyzeData(control):
    """Analyze data callback."""
    if plugin_instance:
        plugin_instance.OnAnalyzeData(control)


def OnTrendAnalysis(control):
    """Trend analysis callback."""
    if plugin_instance:
        plugin_instance.OnTrendAnalysis(control)


def OnPatternDetection(control):
    """Pattern detection callback."""
    if plugin_instance:
        plugin_instance.OnPatternDetection(control)


def OnQueryData(control):
    """Query data callback."""
    if plugin_instance:
        plugin_instance.OnQueryData(control)


def OnQuickQuery(control, text):
    """Quick query callback."""
    if plugin_instance:
        plugin_instance.OnQuickQuery(control, text)


def OnGenerateReport(control):
    """Generate report callback."""
    if plugin_instance:
        plugin_instance.OnGenerateReport(control)


def OnCreateDashboard(control):
    """Create dashboard callback."""
    if plugin_instance:
        plugin_instance.OnCreateDashboard(control)


def OnExportResults(control):
    """Export results callback."""
    if plugin_instance:
        plugin_instance.OnExportResults(control)


def OnModelSelection(control, selectedId, selectedIndex):
    """Model selection callback."""
    if plugin_instance:
        plugin_instance.OnModelSelection(control, selectedId, selectedIndex)


def OnConfigure(control):
    """Configure callback."""
    if plugin_instance:
        plugin_instance.OnConfigure(control)


def OnHelp(control):
    """Help callback."""
    if plugin_instance:
        plugin_instance.OnHelp(control)


def OnRefreshConnection(control):
    """Refresh connection callback."""
    if plugin_instance:
        plugin_instance.OnRefreshConnection(control)


if __name__ == "__main__":
    # Initialize plugin when run directly
    initialize_plugin()
    print("Excel-Ollama AI Plugin initialized successfully!")