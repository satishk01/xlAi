"""
Excel Interface Layer for Excel-Ollama AI Plugin.
Manages all interactions between the plugin and Excel application.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple, Union
import xlwings as xw
import json
import asyncio
from datetime import datetime
import logging

from .interfaces import IExcelDataProvider, IExcelResultWriter, IExcelUIController
from ..utils.config import PluginConfig


class ExcelInterface(IExcelDataProvider, IExcelResultWriter, IExcelUIController):
    """Main interface for Excel COM automation and data exchange."""
    
    def __init__(self, config: PluginConfig):
        self.config = config
        self.app = None
        self.workbook = None
        self.custom_functions = {}
        self.logger = logging.getLogger(__name__)
        
        # Initialize Excel connection
        self._initialize_excel_connection()
        
    def _initialize_excel_connection(self) -> bool:
        """Initialize connection to Excel application."""
        try:
            # Connect to active Excel instance or create new one
            try:
                self.app = xw.apps.active
            except:
                self.app = xw.App(visible=True)
            
            # Get active workbook or create new one
            try:
                self.workbook = self.app.books.active
            except:
                self.workbook = self.app.books.add()
            
            self.logger.info("Excel connection established successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to initialize Excel connection: {e}")
            return False
    
    def register_custom_functions(self) -> bool:
        """Register custom Excel functions (UDFs)."""
        try:
            # Register OLLAMA_ANALYZE function
            @xw.func
            def OLLAMA_ANALYZE(range_ref: str, prompt: str = "Analyze this data") -> str:
                """Custom Excel function for AI analysis."""
                try:
                    # Get data from range
                    data = self.get_range_data(range_ref)
                    if data is None:
                        return "Error: Could not read data from range"
                    
                    # This would typically call the agent controller
                    # For now, return a placeholder
                    return f"Analysis of {data.shape[0]} rows, {data.shape[1]} columns: {prompt}"
                    
                except Exception as e:
                    return f"Error: {str(e)}"
            
            # Register AI_TREND function
            @xw.func
            def AI_TREND(data_range: str, periods: int = 10) -> str:
                """Custom Excel function for trend analysis."""
                try:
                    data = self.get_range_data(data_range)
                    if data is None:
                        return "Error: Could not read data from range"
                    
                    # Placeholder for trend analysis
                    return f"Trend forecast for {periods} periods based on {len(data)} data points"
                    
                except Exception as e:
                    return f"Error: {str(e)}"
            
            # Register PATTERN_DETECT function
            @xw.func
            def PATTERN_DETECT(range_ref: str, threshold: float = 0.5) -> str:
                """Custom Excel function for pattern detection."""
                try:
                    data = self.get_range_data(range_ref)
                    if data is None:
                        return "Error: Could not read data from range"
                    
                    # Placeholder for pattern detection
                    return f"Pattern detection with threshold {threshold} on {data.shape[0]} rows"
                    
                except Exception as e:
                    return f"Error: {str(e)}"
            
            # Store function references
            self.custom_functions = {
                'OLLAMA_ANALYZE': OLLAMA_ANALYZE,
                'AI_TREND': AI_TREND,
                'PATTERN_DETECT': PATTERN_DETECT
            }
            
            self.logger.info("Custom Excel functions registered successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to register custom functions: {e}")
            return False
    
    def get_selected_range(self) -> Optional[pd.DataFrame]:
        """Get data from currently selected Excel range."""
        try:
            if not self.app or not self.workbook:
                return None
            
            # Get the selected range
            selection = self.app.selection
            if not selection:
                return None
            
            # Convert to pandas DataFrame
            values = selection.value
            if not values:
                return None
            
            # Handle single cell
            if not isinstance(values, list):
                return pd.DataFrame([[values]])
            
            # Handle single row
            if not isinstance(values[0], list):
                return pd.DataFrame([values])
            
            # Handle multiple rows and columns
            df = pd.DataFrame(values)
            
            # Try to use first row as headers if they look like strings
            if df.iloc[0].dtype == 'object':
                df.columns = df.iloc[0]
                df = df.drop(df.index[0]).reset_index(drop=True)
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error getting selected range: {e}")
            return None
    
    def get_range_data(self, range_ref: str) -> Optional[pd.DataFrame]:
        """Get data from specified Excel range."""
        try:
            if not self.workbook:
                return None
            
            # Parse range reference
            if '!' in range_ref:
                sheet_name, cell_range = range_ref.split('!', 1)
                sheet = self.workbook.sheets[sheet_name]
            else:
                sheet = self.workbook.sheets.active
                cell_range = range_ref
            
            # Get range values
            range_obj = sheet.range(cell_range)
            values = range_obj.value
            
            if not values:
                return None
            
            # Convert to DataFrame
            if not isinstance(values, list):
                return pd.DataFrame([[values]])
            
            if not isinstance(values[0], list):
                return pd.DataFrame([values])
            
            df = pd.DataFrame(values)
            
            # Try to detect headers
            if len(df) > 1 and df.iloc[0].dtype == 'object':
                df.columns = df.iloc[0]
                df = df.drop(df.index[0]).reset_index(drop=True)
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error getting range data: {e}")
            return None
    
    def get_worksheet_data(self, sheet_name: str = None) -> Optional[pd.DataFrame]:
        """Get all data from specified worksheet."""
        try:
            if not self.workbook:
                return None
            
            if sheet_name:
                sheet = self.workbook.sheets[sheet_name]
            else:
                sheet = self.workbook.sheets.active
            
            # Get used range
            used_range = sheet.used_range
            if not used_range:
                return None
            
            values = used_range.value
            if not values:
                return None
            
            # Convert to DataFrame
            df = pd.DataFrame(values)
            
            # Use first row as headers if appropriate
            if len(df) > 1 and df.iloc[0].dtype == 'object':
                df.columns = df.iloc[0]
                df = df.drop(df.index[0]).reset_index(drop=True)
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error getting worksheet data: {e}")
            return None
    
    def write_results_to_sheet(self, data: Union[pd.DataFrame, Dict, List], 
                              sheet_name: str = "AI_Analysis_Results",
                              start_cell: str = "A1") -> bool:
        """Write analysis results to Excel sheet."""
        try:
            if not self.workbook:
                return False
            
            # Create or get sheet
            try:
                sheet = self.workbook.sheets[sheet_name]
            except:
                sheet = self.workbook.sheets.add(sheet_name)
            
            # Clear existing content
            sheet.clear()
            
            # Handle different data types
            if isinstance(data, pd.DataFrame):
                self._write_dataframe_to_sheet(sheet, data, start_cell)
            elif isinstance(data, dict):
                self._write_dict_to_sheet(sheet, data, start_cell)
            elif isinstance(data, list):
                self._write_list_to_sheet(sheet, data, start_cell)
            else:
                # Convert to string and write
                sheet.range(start_cell).value = str(data)
            
            # Auto-fit columns
            sheet.autofit()
            
            self.logger.info(f"Results written to sheet '{sheet_name}' successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Error writing results to sheet: {e}")
            return False
    
    def create_visualization(self, chart_type: str, data: pd.DataFrame, 
                           sheet_name: str = "AI_Charts") -> bool:
        """Create visualization in Excel."""
        try:
            if not self.workbook or data.empty:
                return False
            
            # Create or get chart sheet
            try:
                sheet = self.workbook.sheets[sheet_name]
            except:
                sheet = self.workbook.sheets.add(sheet_name)
            
            # Write data to sheet first
            data_range = sheet.range('A1').resize(data.shape[0] + 1, data.shape[1])
            
            # Write headers
            sheet.range('A1').resize(1, data.shape[1]).value = list(data.columns)
            
            # Write data
            sheet.range('A2').resize(data.shape[0], data.shape[1]).value = data.values
            
            # Create chart
            chart = sheet.charts.add()
            chart.set_source_data(data_range)
            
            # Set chart type
            chart_types = {
                'line': xw.constants.ChartType.xlLine,
                'bar': xw.constants.ChartType.xlColumnClustered,
                'scatter': xw.constants.ChartType.xlXYScatter,
                'pie': xw.constants.ChartType.xlPie
            }
            
            if chart_type.lower() in chart_types:
                chart.chart_type = chart_types[chart_type.lower()]
            
            # Position chart
            chart.top = sheet.range('A1').top + 200
            chart.left = sheet.range('A1').left
            chart.width = 400
            chart.height = 300
            
            self.logger.info(f"Chart created successfully in sheet '{sheet_name}'")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating visualization: {e}")
            return False
    
    def update_ribbon_status(self, status: str) -> bool:
        """Update status in Excel ribbon."""
        try:
            # This would typically update a custom ribbon control
            # For now, we'll use the status bar
            if self.app:
                self.app.status_bar = f"Ollama AI Plugin: {status}"
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error updating ribbon status: {e}")
            return False
    
    def show_progress_indicator(self, message: str, progress: float = 0) -> bool:
        """Show progress indicator to user."""
        try:
            if self.app:
                status_msg = f"{message} ({progress:.0%})" if progress > 0 else message
                self.app.status_bar = status_msg
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error showing progress indicator: {e}")
            return False
    
    def hide_progress_indicator(self) -> bool:
        """Hide progress indicator."""
        try:
            if self.app:
                self.app.status_bar = "Ready"
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error hiding progress indicator: {e}")
            return False
    
    def get_data_characteristics(self, data: pd.DataFrame) -> Dict[str, Any]:
        """Analyze data characteristics for visualization recommendations."""
        try:
            characteristics = {
                'shape': data.shape,
                'data_types': data.dtypes.to_dict(),
                'numeric_columns': data.select_dtypes(include=[np.number]).columns.tolist(),
                'categorical_columns': data.select_dtypes(include=['object']).columns.tolist(),
                'datetime_columns': data.select_dtypes(include=['datetime64']).columns.tolist(),
                'has_time_series': False,
                'missing_values': data.isnull().sum().sum(),
                'duplicate_rows': data.duplicated().sum()
            }
            
            # Check for time series data
            datetime_cols = characteristics['datetime_columns']
            if len(datetime_cols) > 0:
                characteristics['has_time_series'] = True
                
                # Analyze time series characteristics
                time_col = datetime_cols[0]
                time_series = data[time_col].dropna()
                if len(time_series) > 1:
                    time_diff = time_series.diff().median()
                    characteristics['time_frequency'] = str(time_diff)
                    characteristics['time_range'] = {
                        'start': time_series.min(),
                        'end': time_series.max()
                    }
            
            # Analyze numeric data distribution
            numeric_data = data.select_dtypes(include=[np.number])
            if not numeric_data.empty:
                characteristics['numeric_stats'] = {
                    'mean_values': numeric_data.mean().to_dict(),
                    'std_values': numeric_data.std().to_dict(),
                    'correlation_strength': abs(numeric_data.corr()).mean().mean() if len(numeric_data.columns) > 1 else 0
                }
            
            return characteristics
            
        except Exception as e:
            self.logger.error(f"Error analyzing data characteristics: {e}")
            return {}
    
    def _write_dataframe_to_sheet(self, sheet, data: pd.DataFrame, start_cell: str):
        """Write DataFrame to Excel sheet."""
        # Write headers
        header_range = sheet.range(start_cell).resize(1, len(data.columns))
        header_range.value = list(data.columns)
        header_range.font.bold = True
        
        # Write data
        if not data.empty:
            start_row = sheet.range(start_cell).row + 1
            start_col = sheet.range(start_cell).column
            data_range = sheet.range((start_row, start_col)).resize(len(data), len(data.columns))
            data_range.value = data.values
    
    def _write_dict_to_sheet(self, sheet, data: Dict, start_cell: str):
        """Write dictionary to Excel sheet."""
        row = sheet.range(start_cell).row
        col = sheet.range(start_cell).column
        
        for key, value in data.items():
            # Write key
            sheet.range((row, col)).value = str(key)
            sheet.range((row, col)).font.bold = True
            
            # Write value
            if isinstance(value, (dict, list)):
                sheet.range((row, col + 1)).value = json.dumps(value, indent=2)
            else:
                sheet.range((row, col + 1)).value = str(value)
            
            row += 1
    
    def _write_list_to_sheet(self, sheet, data: List, start_cell: str):
        """Write list to Excel sheet."""
        row = sheet.range(start_cell).row
        col = sheet.range(start_cell).column
        
        for i, item in enumerate(data):
            if isinstance(item, (dict, list)):
                sheet.range((row + i, col)).value = json.dumps(item, indent=2)
            else:
                sheet.range((row + i, col)).value = str(item)
    
    def refresh_data(self) -> bool:
        """Refresh data connections and calculations."""
        try:
            if self.workbook:
                self.workbook.app.calculate()
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error refreshing data: {e}")
            return False
    
    def export_results(self, data: Any, export_format: str = 'xlsx', 
                      filename: str = None) -> bool:
        """Export results to file."""
        try:
            if not filename:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"ollama_analysis_{timestamp}.{export_format}"
            
            if export_format.lower() == 'xlsx':
                # Create new workbook for export
                export_wb = self.app.books.add()
                
                if isinstance(data, pd.DataFrame):
                    self._write_dataframe_to_sheet(export_wb.sheets[0], data, 'A1')
                elif isinstance(data, dict):
                    self._write_dict_to_sheet(export_wb.sheets[0], data, 'A1')
                
                export_wb.save(filename)
                export_wb.close()
                
            elif export_format.lower() == 'csv' and isinstance(data, pd.DataFrame):
                data.to_csv(filename, index=False)
                
            self.logger.info(f"Results exported to {filename}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error exporting results: {e}")
            return False
    
    def cleanup(self):
        """Clean up Excel connections and resources."""
        try:
            if self.app:
                # Don't close the Excel app, just clean up references
                self.app = None
                self.workbook = None
                self.logger.info("Excel interface cleaned up")
                
        except Exception as e:
            self.logger.error(f"Error during cleanup: {e}")


class ExcelFunctionRegistry:
    """Registry for custom Excel functions."""
    
    def __init__(self, excel_interface: ExcelInterface, agent_controller):
        self.excel_interface = excel_interface
        self.agent_controller = agent_controller
        self.registered_functions = {}
    
    async def register_ai_functions(self):
        """Register AI-powered Excel functions."""
        
        @xw.func
        async def OLLAMA_ANALYZE_ASYNC(range_ref: str, prompt: str = "Analyze this data") -> str:
            """Asynchronous AI analysis function."""
            try:
                data = self.excel_interface.get_range_data(range_ref)
                if data is None:
                    return "Error: Could not read data"
                
                # Call agent controller for analysis
                result = await self.agent_controller.execute_analysis_pipeline(
                    data, 'custom_analysis', {'user_query': prompt}
                )
                
                if 'error' in result:
                    return f"Error: {result['error']}"
                
                return result.get('summary', 'Analysis completed')
                
            except Exception as e:
                return f"Error: {str(e)}"
        
        self.registered_functions['OLLAMA_ANALYZE_ASYNC'] = OLLAMA_ANALYZE_ASYNC