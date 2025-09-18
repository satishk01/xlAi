"""
Ollama API client for Excel-Ollama AI Plugin.
Handles communication with Ollama server and model management.
"""

import asyncio
import json
import time
from typing import Dict, List, Optional, Any, AsyncIterator
import aiohttp
import requests
from dataclasses import dataclass
from enum import Enum

from .interfaces import IOllamaClient
from ..utils.config import config_manager


class ModelStatus(Enum):
    """Status of model loading/availability."""
    UNKNOWN = "unknown"
    AVAILABLE = "available"
    LOADING = "loading"
    ERROR = "error"


@dataclass
class ModelInfo:
    """Information about an Ollama model."""
    name: str
    size: int
    modified_at: str
    digest: str
    status: ModelStatus = ModelStatus.UNKNOWN


@dataclass
class GenerationRequest:
    """Request structure for text generation."""
    model: str
    prompt: str
    stream: bool = False
    options: Optional[Dict[str, Any]] = None
    context: Optional[List[int]] = None


class CircuitBreaker:
    """Circuit breaker pattern for handling API failures."""
    
    def __init__(self, failure_threshold: int = 5, recovery_timeout: int = 60):
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.failure_count = 0
        self.last_failure_time = 0
        self.state = "closed"  # closed, open, half-open
    
    def can_execute(self) -> bool:
        """Check if request can be executed."""
        if self.state == "closed":
            return True
        elif self.state == "open":
            if time.time() - self.last_failure_time > self.recovery_timeout:
                self.state = "half-open"
                return True
            return False
        else:  # half-open
            return True
    
    def record_success(self):
        """Record successful request."""
        self.failure_count = 0
        self.state = "closed"
    
    def record_failure(self):
        """Record failed request."""
        self.failure_count += 1
        self.last_failure_time = time.time()
        
        if self.failure_count >= self.failure_threshold:
            self.state = "open"


class OllamaClient(IOllamaClient):
    """Client for communicating with Ollama API."""
    
    def __init__(self, base_url: Optional[str] = None):
        """Initialize Ollama client."""
        config = config_manager.get_config()
        self.base_url = base_url or config.ollama.server_url
        self.timeout = config.ollama.timeout
        self.max_retries = config.ollama.max_retries
        self.stream_responses = config.ollama.stream_responses
        
        self.current_model: Optional[str] = None
        self.model_config: Dict[str, Any] = config.ollama.model_parameters.copy()
        self.available_models: Dict[str, ModelInfo] = {}
        
        # Circuit breaker for handling failures
        self.circuit_breaker = CircuitBreaker()
        
        # Session for async requests
        self._session: Optional[aiohttp.ClientSession] = None
    
    async def _get_session(self) -> aiohttp.ClientSession:
        """Get or create aiohttp session."""
        if self._session is None or self._session.closed:
            timeout = aiohttp.ClientTimeout(total=self.timeout)
            self._session = aiohttp.ClientSession(timeout=timeout)
        return self._session
    
    async def close(self):
        """Close the client session."""
        if self._session and not self._session.closed:
            await self._session.close()
    
    def _make_sync_request(self, method: str, endpoint: str, **kwargs) -> requests.Response:
        """Make synchronous HTTP request with retry logic."""
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        
        for attempt in range(self.max_retries + 1):
            try:
                if not self.circuit_breaker.can_execute():
                    raise Exception("Circuit breaker is open")
                
                response = requests.request(
                    method, url, timeout=self.timeout, **kwargs
                )
                
                if response.status_code == 200:
                    self.circuit_breaker.record_success()
                    return response
                else:
                    response.raise_for_status()
                    
            except Exception as e:
                self.circuit_breaker.record_failure()
                if attempt == self.max_retries:
                    raise e
                
                # Exponential backoff
                wait_time = 2 ** attempt
                time.sleep(wait_time)
        
        raise Exception(f"Failed to make request after {self.max_retries + 1} attempts")
    
    async def _make_async_request(self, method: str, endpoint: str, **kwargs) -> aiohttp.ClientResponse:
        """Make asynchronous HTTP request with retry logic."""
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        session = await self._get_session()
        
        for attempt in range(self.max_retries + 1):
            try:
                if not self.circuit_breaker.can_execute():
                    raise Exception("Circuit breaker is open")
                
                async with session.request(method, url, **kwargs) as response:
                    if response.status == 200:
                        self.circuit_breaker.record_success()
                        return response
                    else:
                        response.raise_for_status()
                        
            except Exception as e:
                self.circuit_breaker.record_failure()
                if attempt == self.max_retries:
                    raise e
                
                # Exponential backoff
                wait_time = 2 ** attempt
                await asyncio.sleep(wait_time)
        
        raise Exception(f"Failed to make request after {self.max_retries + 1} attempts")
    
    def test_connection(self) -> bool:
        """Test connection to Ollama server."""
        try:
            response = self._make_sync_request("GET", "/api/tags")
            return response.status_code == 200
        except Exception:
            return False
    
    async def list_models(self) -> List[str]:
        """Get list of available models."""
        try:
            response = self._make_sync_request("GET", "/api/tags")
            data = response.json()
            
            models = []
            self.available_models.clear()
            
            for model_data in data.get("models", []):
                model_name = model_data["name"]
                models.append(model_name)
                
                self.available_models[model_name] = ModelInfo(
                    name=model_name,
                    size=model_data.get("size", 0),
                    modified_at=model_data.get("modified_at", ""),
                    digest=model_data.get("digest", ""),
                    status=ModelStatus.AVAILABLE
                )
            
            return models
            
        except Exception as e:
            print(f"Error listing models: {e}")
            return []
    
    async def load_model(self, model_name: str) -> bool:
        """Load specified model."""
        try:
            # Check if model is already loaded
            if self.current_model == model_name:
                return True
            
            # Update model status
            if model_name in self.available_models:
                self.available_models[model_name].status = ModelStatus.LOADING
            
            # Make a simple generation request to load the model
            payload = {
                "model": model_name,
                "prompt": "Hello",
                "stream": False,
                "options": {"num_predict": 1}
            }
            
            response = self._make_sync_request("POST", "/api/generate", json=payload)
            
            if response.status_code == 200:
                self.current_model = model_name
                if model_name in self.available_models:
                    self.available_models[model_name].status = ModelStatus.AVAILABLE
                return True
            else:
                if model_name in self.available_models:
                    self.available_models[model_name].status = ModelStatus.ERROR
                return False
                
        except Exception as e:
            print(f"Error loading model {model_name}: {e}")
            if model_name in self.available_models:
                self.available_models[model_name].status = ModelStatus.ERROR
            return False
    
    async def generate_response(self, prompt: str, stream: bool = False) -> str:
        """Generate response from current model."""
        if not self.current_model:
            raise ValueError("No model loaded")
        
        payload = {
            "model": self.current_model,
            "prompt": prompt,
            "stream": stream,
            "options": self.model_config
        }
        
        try:
            if stream:
                return await self._generate_streaming_response(payload)
            else:
                response = self._make_sync_request("POST", "/api/generate", json=payload)
                data = response.json()
                return data.get("response", "")
                
        except Exception as e:
            print(f"Error generating response: {e}")
            raise e
    
    async def _generate_streaming_response(self, payload: Dict[str, Any]) -> str:
        """Generate streaming response."""
        session = await self._get_session()
        url = f"{self.base_url.rstrip('/')}/api/generate"
        
        full_response = ""
        
        try:
            async with session.post(url, json=payload) as response:
                response.raise_for_status()
                
                async for line in response.content:
                    if line:
                        try:
                            data = json.loads(line.decode('utf-8'))
                            if 'response' in data:
                                full_response += data['response']
                            if data.get('done', False):
                                break
                        except json.JSONDecodeError:
                            continue
            
            return full_response
            
        except Exception as e:
            print(f"Error in streaming response: {e}")
            raise e
    
    async def generate_streaming_response(self, prompt: str) -> AsyncIterator[str]:
        """Generate streaming response as async iterator."""
        if not self.current_model:
            raise ValueError("No model loaded")
        
        payload = {
            "model": self.current_model,
            "prompt": prompt,
            "stream": True,
            "options": self.model_config
        }
        
        session = await self._get_session()
        url = f"{self.base_url.rstrip('/')}/api/generate"
        
        try:
            async with session.post(url, json=payload) as response:
                response.raise_for_status()
                
                async for line in response.content:
                    if line:
                        try:
                            data = json.loads(line.decode('utf-8'))
                            if 'response' in data:
                                yield data['response']
                            if data.get('done', False):
                                break
                        except json.JSONDecodeError:
                            continue
                            
        except Exception as e:
            print(f"Error in streaming response: {e}")
            raise e
    
    def configure_model_parameters(self, **kwargs) -> None:
        """Configure model parameters."""
        valid_params = {
            'temperature', 'top_p', 'top_k', 'num_predict', 
            'num_ctx', 'repeat_penalty', 'seed'
        }
        
        for key, value in kwargs.items():
            if key in valid_params:
                self.model_config[key] = value
            else:
                print(f"Warning: Unknown parameter '{key}' ignored")
    
    def get_model_info(self, model_name: str) -> Optional[ModelInfo]:
        """Get information about a specific model."""
        return self.available_models.get(model_name)
    
    def get_current_model(self) -> Optional[str]:
        """Get currently loaded model name."""
        return self.current_model
    
    def get_model_config(self) -> Dict[str, Any]:
        """Get current model configuration."""
        return self.model_config.copy()
    
    async def pull_model(self, model_name: str) -> bool:
        """Pull/download a model from Ollama registry."""
        try:
            payload = {"name": model_name}
            response = self._make_sync_request("POST", "/api/pull", json=payload)
            return response.status_code == 200
        except Exception as e:
            print(f"Error pulling model {model_name}: {e}")
            return False
    
    async def delete_model(self, model_name: str) -> bool:
        """Delete a model from local storage."""
        try:
            payload = {"name": model_name}
            response = self._make_sync_request("DELETE", "/api/delete", json=payload)
            
            if response.status_code == 200:
                # Remove from available models
                if model_name in self.available_models:
                    del self.available_models[model_name]
                
                # Clear current model if it was deleted
                if self.current_model == model_name:
                    self.current_model = None
                
                return True
            return False
            
        except Exception as e:
            print(f"Error deleting model {model_name}: {e}")
            return False
    
    def get_server_info(self) -> Dict[str, Any]:
        """Get Ollama server information."""
        try:
            response = self._make_sync_request("GET", "/api/version")
            return response.json()
        except Exception as e:
            print(f"Error getting server info: {e}")
            return {}
    
    def __del__(self):
        """Cleanup when client is destroyed."""
        if self._session and not self._session.closed:
            # Create a new event loop if none exists
            try:
                loop = asyncio.get_event_loop()
            except RuntimeError:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
            
            if loop.is_running():
                # Schedule cleanup
                loop.create_task(self.close())
            else:
                # Run cleanup
                loop.run_until_complete(self.close())