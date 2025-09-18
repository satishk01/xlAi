"""
Agent Controller for Excel-Ollama AI Plugin.
Orchestrates agent communication and task distribution.
"""

import asyncio
import uuid
import time
from typing import Dict, List, Any, Optional, Type
from dataclasses import dataclass
import pandas as pd

from .interfaces import IOllamaClient
from ..agents.base_agent import BaseAgent, AgentMessage, MessageType, AgentStatus


@dataclass
class AnalysisTask:
    """Represents an analysis task to be processed by agents."""
    task_id: str
    data: pd.DataFrame
    analysis_type: str
    parameters: Dict[str, Any]
    user_query: Optional[str] = None
    priority: int = 1
    created_at: float = None
    
    def __post_init__(self):
        if self.created_at is None:
            self.created_at = time.time()


@dataclass
class AnalysisResult:
    """Result from agent analysis."""
    task_id: str
    agent_id: str
    results: Dict[str, Any]
    confidence_score: float
    methodology: str
    visualizations: List[Dict] = None
    timestamp: float = None
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = time.time()
        if self.visualizations is None:
            self.visualizations = []


class AgentController:
    """Controls and coordinates all agents in the system."""
    
    def __init__(self, ollama_client: IOllamaClient):
        self.ollama_client = ollama_client
        self.agents: Dict[str, BaseAgent] = {}
        self.agent_types: Dict[str, Type[BaseAgent]] = {}
        self.task_queue = asyncio.Queue()
        self.results_cache: Dict[str, AnalysisResult] = {}
        self.message_queue = asyncio.Queue()
        self.running = False
        self._worker_tasks: List[asyncio.Task] = []
    
    def register_agent_type(self, agent_type: str, agent_class: Type[BaseAgent]):
        """Register an agent type for dynamic instantiation."""
        self.agent_types[agent_type] = agent_class
    
    def create_agent(self, agent_type: str, agent_id: Optional[str] = None) -> str:
        """Create and register a new agent instance."""
        if agent_type not in self.agent_types:
            raise ValueError(f"Unknown agent type: {agent_type}")
        
        if agent_id is None:
            agent_id = f"{agent_type}_{uuid.uuid4().hex[:8]}"
        
        agent_class = self.agent_types[agent_type]
        agent = agent_class(agent_id, self.ollama_client)
        self.agents[agent_id] = agent
        
        return agent_id
    
    def get_agent(self, agent_id: str) -> Optional[BaseAgent]:
        """Get agent by ID."""
        return self.agents.get(agent_id)
    
    def get_agents_by_capability(self, analysis_type: str, data_type: str) -> List[BaseAgent]:
        """Get all agents that can handle the given analysis type and data type."""
        capable_agents = []
        for agent in self.agents.values():
            if agent.can_handle(analysis_type, data_type):
                capable_agents.append(agent)
        return capable_agents
    
    async def start(self):
        """Start the agent controller and worker tasks."""
        if self.running:
            return
        
        self.running = True
        
        # Start worker tasks
        self._worker_tasks = [
            asyncio.create_task(self._task_processor()),
            asyncio.create_task(self._message_router())
        ]
    
    async def stop(self):
        """Stop the agent controller and all worker tasks."""
        self.running = False
        
        # Cancel all worker tasks
        for task in self._worker_tasks:
            task.cancel()
        
        # Wait for tasks to complete
        await asyncio.gather(*self._worker_tasks, return_exceptions=True)
        self._worker_tasks.clear()
    
    async def submit_analysis_task(self, task: AnalysisTask) -> str:
        """Submit an analysis task for processing."""
        await self.task_queue.put(task)
        return task.task_id
    
    async def get_analysis_result(self, task_id: str, timeout: float = 30.0) -> Optional[AnalysisResult]:
        """Get analysis result by task ID with timeout."""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            if task_id in self.results_cache:
                return self.results_cache[task_id]
            await asyncio.sleep(0.1)
        
        return None
    
    async def execute_analysis_pipeline(self, data: pd.DataFrame, analysis_type: str, 
                                      parameters: Dict[str, Any] = None) -> AnalysisResult:
        """Execute a complete analysis pipeline."""
        if parameters is None:
            parameters = {}
        
        # Create analysis task
        task = AnalysisTask(
            task_id=str(uuid.uuid4()),
            data=data,
            analysis_type=analysis_type,
            parameters=parameters
        )
        
        # Submit task
        await self.submit_analysis_task(task)
        
        # Wait for result
        result = await self.get_analysis_result(task.task_id)
        if result is None:
            raise TimeoutError(f"Analysis task {task.task_id} timed out")
        
        return result
    
    async def send_message(self, message: AgentMessage):
        """Send a message to the message queue for routing."""
        await self.message_queue.put(message)
    
    async def _task_processor(self):
        """Process tasks from the task queue."""
        while self.running:
            try:
                # Get task from queue with timeout
                task = await asyncio.wait_for(self.task_queue.get(), timeout=1.0)
                
                # Find capable agents
                data_type = self._determine_data_type(task.data)
                capable_agents = self.get_agents_by_capability(task.analysis_type, data_type)
                
                if not capable_agents:
                    # No capable agents found
                    result = AnalysisResult(
                        task_id=task.task_id,
                        agent_id="system",
                        results={"error": f"No agents capable of handling {task.analysis_type} for {data_type}"},
                        confidence_score=0.0,
                        methodology="error"
                    )
                    self.results_cache[task.task_id] = result
                    continue
                
                # Select best agent (for now, just use the first one)
                selected_agent = capable_agents[0]
                
                # Create request message
                request_message = AgentMessage(
                    sender="controller",
                    recipient=selected_agent.agent_id,
                    message_type=MessageType.REQUEST,
                    payload={
                        "task_id": task.task_id,
                        "data": task.data,
                        "analysis_type": task.analysis_type,
                        "parameters": task.parameters,
                        "user_query": task.user_query
                    }
                )
                
                # Process request
                response = await selected_agent.process_message(request_message)
                
                if response and response.message_type == MessageType.RESPONSE:
                    # Create analysis result
                    result = AnalysisResult(
                        task_id=task.task_id,
                        agent_id=selected_agent.agent_id,
                        results=response.payload.get("results", {}),
                        confidence_score=response.payload.get("confidence_score", 0.0),
                        methodology=response.payload.get("methodology", "unknown"),
                        visualizations=response.payload.get("visualizations", [])
                    )
                    self.results_cache[task.task_id] = result
                elif response and response.message_type == MessageType.ERROR:
                    # Handle error
                    result = AnalysisResult(
                        task_id=task.task_id,
                        agent_id=selected_agent.agent_id,
                        results={"error": response.payload.get("error", "Unknown error")},
                        confidence_score=0.0,
                        methodology="error"
                    )
                    self.results_cache[task.task_id] = result
                
            except asyncio.TimeoutError:
                # No tasks in queue, continue
                continue
            except Exception as e:
                # Log error and continue
                print(f"Error processing task: {e}")
                continue
    
    async def _message_router(self):
        """Route messages between agents."""
        while self.running:
            try:
                # Get message from queue with timeout
                message = await asyncio.wait_for(self.message_queue.get(), timeout=1.0)
                
                # Find recipient agent
                recipient_agent = self.agents.get(message.recipient)
                if recipient_agent:
                    # Process message
                    response = await recipient_agent.process_message(message)
                    if response:
                        # Route response back
                        await self.message_queue.put(response)
                
            except asyncio.TimeoutError:
                # No messages in queue, continue
                continue
            except Exception as e:
                # Log error and continue
                print(f"Error routing message: {e}")
                continue
    
    def _determine_data_type(self, data: pd.DataFrame) -> str:
        """Determine the type of data for agent selection."""
        # Simple heuristic - can be made more sophisticated
        if 'date' in data.columns or 'time' in data.columns:
            return "time_series"
        elif len(data.columns) > 10:
            return "multivariate"
        elif data.dtypes.apply(lambda x: x.kind in 'biufc').all():
            return "numerical"
        else:
            return "mixed"
    
    def get_system_status(self) -> Dict[str, Any]:
        """Get overall system status."""
        agent_statuses = {}
        for agent_id, agent in self.agents.items():
            agent_statuses[agent_id] = agent.get_status()
        
        return {
            "running": self.running,
            "agents": agent_statuses,
            "task_queue_size": self.task_queue.qsize(),
            "message_queue_size": self.message_queue.qsize(),
            "results_cache_size": len(self.results_cache)
        }
    
    def clear_cache(self, max_age_seconds: float = 3600):
        """Clear old results from cache."""
        current_time = time.time()
        expired_keys = [
            key for key, result in self.results_cache.items()
            if current_time - result.timestamp > max_age_seconds
        ]
        
        for key in expired_keys:
            del self.results_cache[key]