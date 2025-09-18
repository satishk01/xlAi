"""
Unit tests for Ollama client functionality.
"""

import pytest
import asyncio
from unittest.mock import Mock, patch, AsyncMock
import json

from src.core.ollama_client import OllamaClient, ModelStatus, CircuitBreaker


class TestCircuitBreaker:
    """Test circuit breaker functionality."""
    
    def test_initial_state(self):
        """Test circuit breaker initial state."""
        cb = CircuitBreaker()
        assert cb.can_execute() is True
        assert cb.state == "closed"
    
    def test_failure_threshold(self):
        """Test circuit breaker opens after threshold failures."""
        cb = CircuitBreaker(failure_threshold=3)
        
        # Record failures
        for _ in range(2):
            cb.record_failure()
            assert cb.can_execute() is True
        
        # Third failure should open circuit
        cb.record_failure()
        assert cb.can_execute() is False
        assert cb.state == "open"
    
    def test_recovery(self):
        """Test circuit breaker recovery."""
        cb = CircuitBreaker(failure_threshold=1, recovery_timeout=0.1)
        
        # Trigger circuit open
        cb.record_failure()
        assert cb.can_execute() is False
        
        # Wait for recovery timeout
        import time
        time.sleep(0.2)
        
        # Should be half-open now
        assert cb.can_execute() is True
        assert cb.state == "half-open"
        
        # Success should close circuit
        cb.record_success()
        assert cb.state == "closed"


class TestOllamaClient:
    """Test Ollama client functionality."""
    
    @pytest.fixture
    def client(self):
        """Create test client."""
        return OllamaClient("http://localhost:11434")
    
    @pytest.fixture
    def mock_response(self):
        """Create mock response."""
        mock = Mock()
        mock.status_code = 200
        mock.json.return_value = {"models": []}
        return mock
    
    def test_client_initialization(self, client):
        """Test client initialization."""
        assert client.base_url == "http://localhost:11434"
        assert client.current_model is None
        assert isinstance(client.model_config, dict)
    
    @patch('requests.request')
    def test_test_connection_success(self, mock_request, client):
        """Test successful connection test."""
        mock_request.return_value.status_code = 200
        
        result = client.test_connection()
        assert result is True
        mock_request.assert_called_once()
    
    @patch('requests.request')
    def test_test_connection_failure(self, mock_request, client):
        """Test failed connection test."""
        mock_request.side_effect = Exception("Connection failed")
        
        result = client.test_connection()
        assert result is False
    
    @patch('requests.request')
    def test_list_models(self, mock_request, client):
        """Test listing models."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "models": [
                {
                    "name": "llama2",
                    "size": 1000000,
                    "modified_at": "2023-01-01",
                    "digest": "abc123"
                }
            ]
        }
        mock_request.return_value = mock_response
        
        result = asyncio.run(client.list_models())
        
        assert result == ["llama2"]
        assert "llama2" in client.available_models
        assert client.available_models["llama2"].status == ModelStatus.AVAILABLE
    
    @patch('requests.request')
    def test_load_model_success(self, mock_request, client):
        """Test successful model loading."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"response": "Hello"}
        mock_request.return_value = mock_response
        
        result = asyncio.run(client.load_model("llama2"))
        
        assert result is True
        assert client.current_model == "llama2"
    
    @patch('requests.request')
    def test_load_model_failure(self, mock_request, client):
        """Test failed model loading."""
        mock_request.side_effect = Exception("Model not found")
        
        result = asyncio.run(client.load_model("nonexistent"))
        
        assert result is False
        assert client.current_model is None
    
    @patch('requests.request')
    def test_generate_response(self, mock_request, client):
        """Test response generation."""
        # First call to load model
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"response": "Hello"}
        mock_request.return_value = mock_response
        
        # Load model first
        asyncio.run(client.load_model("llama2"))
        
        # Test generation
        mock_response.json.return_value = {"response": "Test response"}
        result = asyncio.run(client.generate_response("Test prompt"))
        
        assert result == "Test response"
    
    def test_generate_response_no_model(self, client):
        """Test response generation without loaded model."""
        with pytest.raises(ValueError, match="No model loaded"):
            asyncio.run(client.generate_response("Test prompt"))
    
    def test_configure_model_parameters(self, client):
        """Test model parameter configuration."""
        client.configure_model_parameters(
            temperature=0.8,
            top_p=0.9,
            invalid_param="should_be_ignored"
        )
        
        assert client.model_config["temperature"] == 0.8
        assert client.model_config["top_p"] == 0.9
        assert "invalid_param" not in client.model_config
    
    def test_get_model_info(self, client):
        """Test getting model information."""
        # Add a model to available models
        from src.core.ollama_client import ModelInfo
        client.available_models["test_model"] = ModelInfo(
            name="test_model",
            size=1000,
            modified_at="2023-01-01",
            digest="abc123"
        )
        
        info = client.get_model_info("test_model")
        assert info is not None
        assert info.name == "test_model"
        assert info.size == 1000
        
        # Test non-existent model
        info = client.get_model_info("nonexistent")
        assert info is None
    
    def test_get_current_model(self, client):
        """Test getting current model."""
        assert client.get_current_model() is None
        
        client.current_model = "llama2"
        assert client.get_current_model() == "llama2"
    
    def test_get_model_config(self, client):
        """Test getting model configuration."""
        config = client.get_model_config()
        assert isinstance(config, dict)
        
        # Modify original config
        client.model_config["temperature"] = 0.5
        
        # Returned config should be a copy
        new_config = client.get_model_config()
        new_config["temperature"] = 0.8
        
        assert client.model_config["temperature"] == 0.5
    
    @patch('requests.request')
    def test_pull_model(self, mock_request, client):
        """Test pulling a model."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_request.return_value = mock_response
        
        result = asyncio.run(client.pull_model("llama2"))
        assert result is True
    
    @patch('requests.request')
    def test_delete_model(self, mock_request, client):
        """Test deleting a model."""
        # Add model to available models
        from src.core.ollama_client import ModelInfo
        client.available_models["test_model"] = ModelInfo(
            name="test_model",
            size=1000,
            modified_at="2023-01-01",
            digest="abc123"
        )
        client.current_model = "test_model"
        
        mock_response = Mock()
        mock_response.status_code = 200
        mock_request.return_value = mock_response
        
        result = asyncio.run(client.delete_model("test_model"))
        
        assert result is True
        assert "test_model" not in client.available_models
        assert client.current_model is None
    
    @patch('requests.request')
    def test_get_server_info(self, mock_request, client):
        """Test getting server information."""
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"version": "0.1.0"}
        mock_request.return_value = mock_response
        
        info = client.get_server_info()
        assert info == {"version": "0.1.0"}
    
    @patch('requests.request')
    def test_retry_mechanism(self, mock_request, client):
        """Test retry mechanism on failures."""
        # First two calls fail, third succeeds
        mock_request.side_effect = [
            Exception("Connection failed"),
            Exception("Connection failed"),
            Mock(status_code=200, json=lambda: {"models": []})
        ]
        
        result = asyncio.run(client.list_models())
        assert result == []
        assert mock_request.call_count == 3
    
    @patch('requests.request')
    def test_circuit_breaker_integration(self, mock_request, client):
        """Test circuit breaker integration."""
        # Configure circuit breaker with low threshold
        client.circuit_breaker = CircuitBreaker(failure_threshold=2)
        
        # Cause failures to open circuit
        mock_request.side_effect = Exception("Connection failed")
        
        # First two attempts should fail and open circuit
        with pytest.raises(Exception):
            asyncio.run(client.list_models())
        
        # Circuit should be open now, preventing further requests
        assert client.circuit_breaker.state == "open"
        
        # Next request should fail immediately due to open circuit
        with pytest.raises(Exception, match="Circuit breaker is open"):
            asyncio.run(client.list_models())


if __name__ == "__main__":
    pytest.main([__file__])