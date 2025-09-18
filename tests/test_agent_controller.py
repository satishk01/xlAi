"""
Unit tests for Agent Controller.
"""

import pytest
import asyncio
import pandas as pd
from unittest.mock import Mock, AsyncMock
import sys
import os

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from core.agent_controller import AgentController, AnalysisTask, AnalysisResult
from core.interfaces import IOllamaClient
from agents.base_agent import BaseAgent, AgentCapability, AgentMessage, MessageType


class MockAgent(BaseAgent):
    """Mock agent for testing."""
    
    def _define_capabilities(self) -> AgentCapability:
        return AgentCapability(
            agent_type="mock",
            supported_analysis_types=["test_analysis"],
            required_data_types=["numerical"],
            output_formats=["json"]
        )
    
    async def _process_request(self, payload):
        return {
            "results": {"test": "success"},
            "confidence_score": 0.95,
            "methodology": "mock_analysis"
        }


@pytest.fixture
def mock_ollama_client():
    """Create mock Ollama client."""
    client = Mock(spec=IOllamaClient)
    client.generate_response = AsyncMock(return_value="Mock response")
    return client


@pytest.fixture
def agent_controller(mock_ollama_client):
    """Create agent controller with mock client."""
    return AgentController(mock_ollama_client)


@pytest.fixture
def sample_data():
    """Create sample data for testing."""
    return pd.DataFrame({
        'value1': [1, 2, 3, 4, 5],
        'value2': [10, 20, 30, 40, 50]
    })


@pytest.mark.asyncio
async def test_agent_registration(agent_controller):
    """Test agent type registration and creation."""
    # Register agent type
    agent_controller.register_agent_type("mock", MockAgent)
    
    # Create agent
    agent_id = agent_controller.create_agent("mock")
    
    # Verify agent was created
    assert agent_id in agent_controller.agents
    agent = agent_controller.get_agent(agent_id)
    assert isinstance(agent, MockAgent)
    assert agent.agent_id == agent_id


@pytest.mark.asyncio
async def test_agent_capability_matching(agent_controller):
    """Test finding agents by capability."""
    # Register and create agent
    agent_controller.register_agent_type("mock", MockAgent)
    agent_id = agent_controller.create_agent("mock")
    
    # Test capability matching
    capable_agents = agent_controller.get_agents_by_capability("test_analysis", "numerical")
    assert len(capable_agents) == 1
    assert capable_agents[0].agent_id == agent_id
    
    # Test no match
    no_match = agent_controller.get_agents_by_capability("unknown_analysis", "numerical")
    assert len(no_match) == 0


@pytest.mark.asyncio
async def test_task_submission_and_processing(agent_controller, sample_data):
    """Test task submission and processing."""
    # Register and create agent
    agent_controller.register_agent_type("mock", MockAgent)
    agent_controller.create_agent("mock")
    
    # Start controller
    await agent_controller.start()
    
    try:
        # Create and submit task
        task = AnalysisTask(
            task_id="test_task",
            data=sample_data,
            analysis_type="test_analysis",
            parameters={"param1": "value1"}
        )
        
        task_id = await agent_controller.submit_analysis_task(task)
        assert task_id == "test_task"
        
        # Wait for result
        result = await agent_controller.get_analysis_result(task_id, timeout=5.0)
        
        # Verify result
        assert result is not None
        assert result.task_id == task_id
        assert result.results["test"] == "success"
        assert result.confidence_score == 0.95
        
    finally:
        await agent_controller.stop()


@pytest.mark.asyncio
async def test_analysis_pipeline(agent_controller, sample_data):
    """Test complete analysis pipeline."""
    # Register and create agent
    agent_controller.register_agent_type("mock", MockAgent)
    agent_controller.create_agent("mock")
    
    # Start controller
    await agent_controller.start()
    
    try:
        # Execute pipeline
        result = await agent_controller.execute_analysis_pipeline(
            data=sample_data,
            analysis_type="test_analysis",
            parameters={"test": True}
        )
        
        # Verify result
        assert result.results["test"] == "success"
        assert result.confidence_score == 0.95
        assert result.methodology == "mock_analysis"
        
    finally:
        await agent_controller.stop()


@pytest.mark.asyncio
async def test_no_capable_agents(agent_controller, sample_data):
    """Test handling when no agents can handle the task."""
    # Start controller without registering any agents
    await agent_controller.start()
    
    try:
        # Create task that no agent can handle
        task = AnalysisTask(
            task_id="impossible_task",
            data=sample_data,
            analysis_type="impossible_analysis",
            parameters={}
        )
        
        await agent_controller.submit_analysis_task(task)
        result = await agent_controller.get_analysis_result("impossible_task", timeout=2.0)
        
        # Should get error result
        assert result is not None
        assert "error" in result.results
        assert result.confidence_score == 0.0
        
    finally:
        await agent_controller.stop()


@pytest.mark.asyncio
async def test_message_routing(agent_controller):
    """Test message routing between agents."""
    # Register and create agent
    agent_controller.register_agent_type("mock", MockAgent)
    agent_id = agent_controller.create_agent("mock")
    
    # Start controller
    await agent_controller.start()
    
    try:
        # Create test message
        message = AgentMessage(
            sender="test_sender",
            recipient=agent_id,
            message_type=MessageType.REQUEST,
            payload={"test": "data"}
        )
        
        # Send message
        await agent_controller.send_message(message)
        
        # Give some time for processing
        await asyncio.sleep(0.1)
        
        # Message should be processed (no exceptions thrown)
        assert True
        
    finally:
        await agent_controller.stop()


def test_data_type_determination(agent_controller):
    """Test data type determination logic."""
    # Time series data
    time_data = pd.DataFrame({
        'date': pd.date_range('2023-01-01', periods=5),
        'value': [1, 2, 3, 4, 5]
    })
    assert agent_controller._determine_data_type(time_data) == "time_series"
    
    # Multivariate data
    multi_data = pd.DataFrame({f'col_{i}': range(5) for i in range(15)})
    assert agent_controller._determine_data_type(multi_data) == "multivariate"
    
    # Numerical data
    num_data = pd.DataFrame({'a': [1, 2, 3], 'b': [4.0, 5.0, 6.0]})
    assert agent_controller._determine_data_type(num_data) == "numerical"
    
    # Mixed data
    mixed_data = pd.DataFrame({'num': [1, 2, 3], 'text': ['a', 'b', 'c']})
    assert agent_controller._determine_data_type(mixed_data) == "mixed"


def test_system_status(agent_controller):
    """Test system status reporting."""
    # Register and create agent
    agent_controller.register_agent_type("mock", MockAgent)
    agent_id = agent_controller.create_agent("mock")
    
    # Get status
    status = agent_controller.get_system_status()
    
    # Verify status structure
    assert "running" in status
    assert "agents" in status
    assert "task_queue_size" in status
    assert "message_queue_size" in status
    assert "results_cache_size" in status
    
    # Verify agent status
    assert agent_id in status["agents"]
    assert status["agents"][agent_id]["agent_id"] == agent_id


def test_cache_clearing(agent_controller):
    """Test results cache clearing."""
    # Add some results to cache
    old_result = AnalysisResult(
        task_id="old_task",
        agent_id="test_agent",
        results={"test": "old"},
        confidence_score=0.5,
        methodology="test"
    )
    old_result.timestamp = 0  # Very old timestamp
    
    recent_result = AnalysisResult(
        task_id="recent_task",
        agent_id="test_agent",
        results={"test": "recent"},
        confidence_score=0.8,
        methodology="test"
    )
    
    agent_controller.results_cache["old_task"] = old_result
    agent_controller.results_cache["recent_task"] = recent_result
    
    # Clear old results
    agent_controller.clear_cache(max_age_seconds=3600)
    
    # Old result should be removed, recent should remain
    assert "old_task" not in agent_controller.results_cache
    assert "recent_task" in agent_controller.results_cache


if __name__ == "__main__":
    pytest.main([__file__])