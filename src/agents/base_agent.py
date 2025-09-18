"""
Base agent class for Excel-Ollama AI Plugin.
Provides common functionality for all specialized agents.
"""

import asyncio
import uuid
import time
from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional, Callable
from dataclasses import dataclass, field
from enum import Enum
import pandas as pd

from ..core.interfaces import IOllamaClient


class MessageType(Enum):
    """Types of messages between agents."""
    REQUEST = "request"
    RESPONSE = "response"
    NOTIFICATION = "notification"
    ERROR = "error"


class AgentStatus(Enum):
    """Status of agent execution."""
    IDLE = "idle"
    BUSY = "busy"
    ERROR = "error"
    STOPPED = "stopped"


@dataclass
class AgentMessage:
    """Message structure for inter-agent communication."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    sender: str = ""
    recipient: str = ""
    message_type: MessageType = MessageType.REQUEST
    payload: Dict[str, Any] = field(default_factory=dict)
    timestamp: float = field(default_factory=time.time)
    correlation_id: Optional[str] = None


@dataclass
class AgentCapability:
    """Defines what an agent can do."""
    agent_type: str
    supported_analysis_types: List[str]
    required_data_types: List[str]
    output_formats: List[str]


class BaseAgent(ABC):
    """Base class for all AI agents in the system."""
    
    def __init__(self, agent_id: str, ollama_client: IOllamaClient):
        self.agent_id = agent_id
        self.ollama_client = ollama_client
        self.status = AgentStatus.IDLE
        self.capabilities = self._define_capabilities()
        self.message_handlers: Dict[MessageType, Callable] = {}
        self.results_cache: Dict[str, Any] = {}
        self._setup_message_handlers()
    
    @abstractmethod
    def _define_capabilities(self) -> AgentCapability:
        """Define what this agent can do."""
        pass
    
    def _setup_message_handlers(self):
        """Set up message handlers for different message types."""
        self.message_handlers = {
            MessageType.REQUEST: self._handle_request,
            MessageType.RESPONSE: self._handle_response,
            MessageType.NOTIFICATION: self._handle_notification,
            MessageType.ERROR: self._handle_error
        }
    
    async def process_message(self, message: AgentMessage) -> Optional[AgentMessage]:
        """Process incoming message and return response if needed."""
        try:
            handler = self.message_handlers.get(message.message_type)
            if handler:
                return await handler(message)
            else:
                return AgentMessage(
                    sender=self.agent_id,
                    recipient=message.sender,
                    message_type=MessageType.ERROR,
                    payload={"error": f"Unknown message type: {message.message_type}"},
                    correlation_id=message.id
                )
        except Exception as e:
            return AgentMessage(
                sender=self.agent_id,
                recipient=message.sender,
                message_type=MessageType.ERROR,
                payload={"error": str(e)},
                correlation_id=message.id
            )
    
    async def _handle_request(self, message: AgentMessage) -> AgentMessage:
        """Handle incoming requests."""
        self.status = AgentStatus.BUSY
        try:
            result = await self._process_request(message.payload)
            self.status = AgentStatus.IDLE
            return AgentMessage(
                sender=self.agent_id,
                recipient=message.sender,
                message_type=MessageType.RESPONSE,
                payload=result,
                correlation_id=message.id
            )
        except Exception as e:
            self.status = AgentStatus.ERROR
            return AgentMessage(
                sender=self.agent_id,
                recipient=message.sender,
                message_type=MessageType.ERROR,
                payload={"error": str(e)},
                correlation_id=message.id
            )
    
    async def _handle_response(self, message: AgentMessage) -> None:
        """Handle responses from other agents."""
        # Store response in cache for correlation
        if message.correlation_id:
            self.results_cache[message.correlation_id] = message.payload
    
    async def _handle_notification(self, message: AgentMessage) -> None:
        """Handle notifications from other agents."""
        # Default implementation - can be overridden by subclasses
        pass
    
    async def _handle_error(self, message: AgentMessage) -> None:
        """Handle error messages from other agents."""
        # Default implementation - can be overridden by subclasses
        pass
    
    @abstractmethod
    async def _process_request(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Process the actual request. Must be implemented by subclasses."""
        pass
    
    def can_handle(self, analysis_type: str, data_type: str) -> bool:
        """Check if this agent can handle the given analysis type and data type."""
        return (analysis_type in self.capabilities.supported_analysis_types and
                data_type in self.capabilities.required_data_types)
    
    async def generate_llm_response(self, prompt: str, **kwargs) -> str:
        """Generate response using Ollama LLM."""
        try:
            response = await self.ollama_client.generate_response(prompt, **kwargs)
            return response
        except Exception as e:
            raise Exception(f"LLM generation failed: {str(e)}")
    
    def get_status(self) -> Dict[str, Any]:
        """Get current agent status and metrics."""
        return {
            "agent_id": self.agent_id,
            "status": self.status.value,
            "capabilities": self.capabilities,
            "cache_size": len(self.results_cache)
        }