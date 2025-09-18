"""
Core interfaces for the Excel-Ollama AI Plugin.
Defines abstract base classes and protocols for system components.
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional, Iterator
import pandas as pd
from dataclasses import dataclass
from datetime import datetime
from enum import Enum


class MessageType(Enum):
    """Types of messages exchanged between agents."""
    REQUEST = "request"
    RESPONSE = "response"
    NOTIFICATION = "notification"


@dataclass
class AgentMessage:
    """Message structure for inter-agent communication."""
    sender: str
    recipient: str
    message_type: MessageType
    payload: Dict[str, Any]
    timestamp: datetime
    message_id: str


@dataclass
class AnalysisRequest:
    """Request structure for data analysis operations."""
    data: pd.DataFrame
    analysis_type: str
    parameters: Dict[str, Any]
    user_query: Optional[str] = None
    model_preference: Optional[str] = None
    request_id: str = ""


@dataclass
class AnalysisResult:
    """Result structure for completed analysis operations."""
    request_id: str
    agent_type: str
    results: Dict[str, Any]
    confidence_score: float
    methodology: str
    visualizations: List[Dict]
    timestamp: datetime


@dataclass
class ValidationResult:
    """Result of data validation operations."""
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    data_types: Dict[str, str]
    row_count: int
    column_count: int


class IExcelDataProvider(ABC):
    """Interface for reading data from Excel."""
    
    @abstractmethod
    def get_selected_range(self) -> pd.DataFrame:
        """Get data from currently selected Excel range."""
        pass
    
    @abstractmethod
    def get_sheet_data(self, sheet_name: str) -> pd.DataFrame:
        """Get all data from specified sheet."""
        pass
    
    @abstractmethod
    def get_range_data(self, range_address: str) -> pd.DataFrame:
        """Get data from specific range address."""
        pass


class IExcelResultWriter(ABC):
    """Interface for writing results to Excel."""
    
    @abstractmethod
    def write_results_to_sheet(self, data: pd.DataFrame, sheet_name: str) -> bool:
        """Write analysis results to Excel sheet."""
        pass
    
    @abstractmethod
    def create_visualization(self, chart_type: str, data: Dict[str, Any]) -> bool:
        """Create chart/visualization in Excel."""
        pass
    
    @abstractmethod
    def update_cell_formula(self, cell_address: str, formula: str) -> bool:
        """Update Excel cell with custom formula."""
        pass


class IExcelUIController(ABC):
    """Interface for Excel UI interactions."""
    
    @abstractmethod
    def update_ribbon_status(self, status: str) -> None:
        """Update status display in Excel ribbon."""
        pass
    
    @abstractmethod
    def show_progress(self, percentage: float, message: str) -> None:
        """Show progress indicator."""
        pass
    
    @abstractmethod
    def show_dialog(self, dialog_type: str, **kwargs) -> Dict[str, Any]:
        """Display dialog and return user input."""
        pass


class IDataProcessor(ABC):
    """Interface for data processing operations."""
    
    @abstractmethod
    def validate_data(self, data: pd.DataFrame) -> ValidationResult:
        """Validate data quality and structure."""
        pass
    
    @abstractmethod
    def clean_data(self, data: pd.DataFrame, cleaning_rules: Dict[str, Any]) -> pd.DataFrame:
        """Clean and preprocess data."""
        pass
    
    @abstractmethod
    def chunk_large_dataset(self, data: pd.DataFrame) -> Iterator[pd.DataFrame]:
        """Split large dataset into manageable chunks."""
        pass


class IOllamaClient(ABC):
    """Interface for Ollama API communication."""
    
    @abstractmethod
    async def list_models(self) -> List[str]:
        """Get list of available models."""
        pass
    
    @abstractmethod
    async def load_model(self, model_name: str) -> bool:
        """Load specified model."""
        pass
    
    @abstractmethod
    async def generate_response(self, prompt: str, stream: bool = False) -> str:
        """Generate response from model."""
        pass
    
    @abstractmethod
    def configure_model_parameters(self, **kwargs) -> None:
        """Configure model parameters."""
        pass


class IAgent(ABC):
    """Base interface for all AI agents."""
    
    @abstractmethod
    async def process_request(self, request: AnalysisRequest) -> AnalysisResult:
        """Process analysis request and return results."""
        pass
    
    @abstractmethod
    async def handle_message(self, message: AgentMessage) -> Optional[AgentMessage]:
        """Handle inter-agent communication."""
        pass
    
    @abstractmethod
    def get_capabilities(self) -> Dict[str, Any]:
        """Return agent capabilities and supported operations."""
        pass


class IAgentController(ABC):
    """Interface for agent coordination and management."""
    
    @abstractmethod
    async def execute_analysis_pipeline(self, data: pd.DataFrame, analysis_type: str) -> AnalysisResult:
        """Execute complete analysis pipeline."""
        pass
    
    @abstractmethod
    async def coordinate_agents(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """Coordinate multiple agents for complex tasks."""
        pass
    
    @abstractmethod
    def register_agent(self, agent: IAgent) -> bool:
        """Register new agent with controller."""
        pass