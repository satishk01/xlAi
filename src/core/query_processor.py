"""
Natural Language Query Processor for Excel-Ollama AI Plugin.
Processes user queries and converts them to analysis operations.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple
import re
import json
import asyncio
from datetime import datetime
import logging

from .ollama_client import OllamaClient


class QueryProcessor:
    """Processes natural language queries and converts them to analysis operations."""
    
    def __init__(self, ollama_client: OllamaClient):
        self.ollama_client = ollama_client
        self.logger = logging.getLogger(__name__)
        
        # Query patterns and mappings
        self.query_patterns = self._initialize_query_patterns()
        self.analysis_mappings = self._initialize_analysis_mappings()
        
    def _initialize_query_patterns(self) -> Dict[str, List[str]]:
        """Initialize common query patterns for intent recognition."""
        return {
            'trend_analysis': [
                r'trend', r'trending', r'increase', r'decrease', r'growing', r'declining',
                r'over time', r'time series', r'forecast', r'predict', r'future'
            ],
            'statistical_analysis': [
                r'average', r'mean', r'median', r'standard deviation', r'variance',
                r'statistics', r'summary', r'describe', r'distribution'
            ],
            'correlation_analysis': [
                r'correlation', r'relationship', r'related', r'connected', r'associated',
                r'depends on', r'affects', r'influence'
            ],
            'pattern_detection': [
                r'pattern', r'seasonal', r'cyclical', r'recurring', r'regular',
                r'anomaly', r'outlier', r'unusual', r'abnormal'
            ],
            'comparison': [
                r'compare', r'comparison', r'versus', r'vs', r'difference', r'between',
                r'higher', r'lower', r'better', r'worse'
            ],
            'clustering': [
                r'group', r'cluster', r'segment', r'category', r'similar', r'alike',
                r'classify', r'categorize'
            ],
            'forecasting': [
                r'forecast', r'predict', r'future', r'next', r'upcoming', r'projection',
                r'estimate', r'expect'
            ]
        }
    
    def _initialize_analysis_mappings(self) -> Dict[str, Dict[str, Any]]:
        """Initialize mappings from query types to analysis operations."""
        return {
            'trend_analysis': {
                'agent': 'analysis',
                'method': 'analyze_trends',
                'requires': ['time_column', 'value_columns'],
                'optional': ['period', 'method']
            },
            'statistical_analysis': {
                'agent': 'analysis',
                'method': 'calculate_statistics',
                'requires': ['data'],
                'optional': ['columns']
            },
            'correlation_analysis': {
                'agent': 'analysis',
                'method': 'calculate_statistics',
                'requires': ['data'],
                'optional': ['columns'],
                'focus': 'correlations'
            },
            'pattern_detection': {
                'agent': 'pattern',
                'method': 'detect_seasonal_patterns',
                'requires': ['data'],
                'optional': ['frequency']
            },
            'clustering': {
                'agent': 'pattern',
                'method': 'cluster_activities',
                'requires': ['data'],
                'optional': ['features', 'n_clusters']
            },
            'forecasting': {
                'agent': 'analysis',
                'method': 'perform_forecasting',
                'requires': ['data'],
                'optional': ['periods', 'method']
            }
        }
    
    async def process_query(self, query: str, data: pd.DataFrame, 
                          context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Process natural language query and return analysis specification."""
        try:
            # Clean and normalize query
            normalized_query = self._normalize_query(query)
            
            # Detect query intent
            intent = await self._detect_intent(normalized_query, data)
            
            # Extract parameters from query
            parameters = await self._extract_parameters(normalized_query, data, intent)
            
            # Generate clarifying questions if needed
            clarifications = await self._generate_clarifications(intent, parameters, data)
            
            # Create analysis specification
            analysis_spec = self._create_analysis_specification(intent, parameters, data)
            
            # Generate explanation
            explanation = await self._generate_explanation(query, analysis_spec)
            
            return {
                'original_query': query,
                'normalized_query': normalized_query,
                'detected_intent': intent,
                'parameters': parameters,
                'clarifications': clarifications,
                'analysis_specification': analysis_spec,
                'explanation': explanation,
                'confidence_score': self._calculate_query_confidence(intent, parameters),
                'processing_timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            self.logger.error(f"Error processing query: {e}")
            return {
                'error': str(e),
                'original_query': query,
                'processing_timestamp': datetime.now().isoformat()
            }
    
    async def _detect_intent(self, query: str, data: pd.DataFrame) -> Dict[str, Any]:
        """Detect the intent of the user query."""
        # Rule-based intent detection
        intent_scores = {}
        
        for intent_type, patterns in self.query_patterns.items():
            score = 0
            for pattern in patterns:
                matches = len(re.findall(pattern, query, re.IGNORECASE))
                score += matches
            
            if score > 0:
                intent_scores[intent_type] = score
        
        # LLM-based intent detection for complex queries
        llm_intent = await self._llm_intent_detection(query, data)
        
        # Combine rule-based and LLM results
        primary_intent = max(intent_scores.items(), key=lambda x: x[1])[0] if intent_scores else 'general_analysis'
        
        return {
            'primary_intent': primary_intent,
            'intent_scores': intent_scores,
            'llm_intent': llm_intent,
            'confidence': max(intent_scores.values()) / len(query.split()) if intent_scores else 0.3
        }
    
    async def _llm_intent_detection(self, query: str, data: pd.DataFrame) -> Dict[str, Any]:
        """Use LLM to detect query intent for complex cases."""
        try:
            data_info = self._get_data_summary(data)
            
            prompt = f"""
            Analyze this user query about data analysis and determine the intent:
            
            Query: "{query}"
            
            Data Information:
            - Shape: {data.shape}
            - Columns: {list(data.columns)}
            - Data types: {data.dtypes.to_dict()}
            
            Available analysis types:
            - trend_analysis: Analyze trends over time
            - statistical_analysis: Calculate descriptive statistics
            - correlation_analysis: Find relationships between variables
            - pattern_detection: Detect patterns and anomalies
            - clustering: Group similar data points
            - forecasting: Predict future values
            - comparison: Compare different groups or time periods
            
            Respond with:
            1. Primary intent (one of the analysis types above)
            2. Secondary intents (if any)
            3. Confidence level (0-1)
            4. Reasoning
            
            Format as JSON.
            """
            
            response = await self.ollama_client.generate_response(prompt)
            
            # Try to parse JSON response
            try:
                return json.loads(response)
            except:
                # Fallback to text parsing
                return {
                    'primary_intent': 'general_analysis',
                    'confidence': 0.5,
                    'reasoning': response
                }
                
        except Exception as e:
            self.logger.error(f"Error in LLM intent detection: {e}")
            return {
                'primary_intent': 'general_analysis',
                'confidence': 0.3,
                'error': str(e)
            }
    
    async def _extract_parameters(self, query: str, data: pd.DataFrame, 
                                intent: Dict[str, Any]) -> Dict[str, Any]:
        """Extract parameters from the query based on detected intent."""
        parameters = {}
        
        primary_intent = intent['primary_intent']
        
        # Extract column references
        columns = self._extract_column_references(query, data)
        if columns:
            parameters['columns'] = columns
        
        # Extract time-related parameters
        time_params = self._extract_time_parameters(query, data)
        parameters.update(time_params)
        
        # Extract numeric parameters
        numeric_params = self._extract_numeric_parameters(query)
        parameters.update(numeric_params)
        
        # Intent-specific parameter extraction
        if primary_intent == 'trend_analysis':
            parameters.update(self._extract_trend_parameters(query, data))
        elif primary_intent == 'forecasting':
            parameters.update(self._extract_forecast_parameters(query))
        elif primary_intent == 'clustering':
            parameters.update(self._extract_clustering_parameters(query))
        elif primary_intent == 'comparison':
            parameters.update(self._extract_comparison_parameters(query, data))
        
        # Use LLM for complex parameter extraction
        llm_params = await self._llm_parameter_extraction(query, data, primary_intent)
        parameters.update(llm_params)
        
        return parameters
    
    def _extract_column_references(self, query: str, data: pd.DataFrame) -> List[str]:
        """Extract column references from the query."""
        referenced_columns = []
        
        # Look for exact column name matches
        for column in data.columns:
            if column.lower() in query.lower():
                referenced_columns.append(column)
        
        # Look for partial matches
        query_words = query.lower().split()
        for column in data.columns:
            column_words = column.lower().split('_')
            if any(word in query_words for word in column_words):
                if column not in referenced_columns:
                    referenced_columns.append(column)
        
        return referenced_columns
    
    def _extract_time_parameters(self, query: str, data: pd.DataFrame) -> Dict[str, Any]:
        """Extract time-related parameters from the query."""
        params = {}
        
        # Find datetime columns
        datetime_columns = data.select_dtypes(include=['datetime64']).columns.tolist()
        if datetime_columns:
            params['time_column'] = datetime_columns[0]
        
        # Extract time periods
        period_patterns = {
            'daily': r'daily|day|days',
            'weekly': r'weekly|week|weeks',
            'monthly': r'monthly|month|months',
            'yearly': r'yearly|year|years|annual'
        }
        
        for period, pattern in period_patterns.items():
            if re.search(pattern, query, re.IGNORECASE):
                params['frequency'] = period
                break
        
        return params
    
    def _extract_numeric_parameters(self, query: str) -> Dict[str, Any]:
        """Extract numeric parameters from the query."""
        params = {}
        
        # Extract numbers
        numbers = re.findall(r'\b\d+\.?\d*\b', query)
        
        # Context-based number interpretation
        if 'forecast' in query.lower() or 'predict' in query.lower():
            if numbers:
                params['periods'] = int(float(numbers[0]))
        
        if 'cluster' in query.lower() or 'group' in query.lower():
            if numbers:
                params['n_clusters'] = int(float(numbers[0]))
        
        if 'threshold' in query.lower():
            if numbers:
                params['threshold'] = float(numbers[0])
        
        return params
    
    def _extract_trend_parameters(self, query: str, data: pd.DataFrame) -> Dict[str, Any]:
        """Extract trend analysis specific parameters."""
        params = {}
        
        # Identify value columns for trend analysis
        numeric_columns = data.select_dtypes(include=[np.number]).columns.tolist()
        if numeric_columns:
            params['value_columns'] = numeric_columns
        
        # Extract trend direction interest
        if any(word in query.lower() for word in ['increase', 'growing', 'rising']):
            params['trend_direction'] = 'increasing'
        elif any(word in query.lower() for word in ['decrease', 'declining', 'falling']):
            params['trend_direction'] = 'decreasing'
        
        return params
    
    def _extract_forecast_parameters(self, query: str) -> Dict[str, Any]:
        """Extract forecasting specific parameters."""
        params = {}
        
        # Default forecast periods
        params['periods'] = 10
        
        # Look for specific time horizons
        if 'next month' in query.lower():
            params['periods'] = 30
        elif 'next week' in query.lower():
            params['periods'] = 7
        elif 'next year' in query.lower():
            params['periods'] = 365
        
        return params
    
    def _extract_clustering_parameters(self, query: str) -> Dict[str, Any]:
        """Extract clustering specific parameters."""
        params = {}
        
        # Default number of clusters
        params['n_clusters'] = 3
        
        # Look for grouping hints
        group_words = ['group', 'cluster', 'segment', 'category']
        for word in group_words:
            if word in query.lower():
                # Look for numbers near grouping words
                pattern = rf'{word}\s*(\d+)'
                match = re.search(pattern, query, re.IGNORECASE)
                if match:
                    params['n_clusters'] = int(match.group(1))
        
        return params
    
    def _extract_comparison_parameters(self, query: str, data: pd.DataFrame) -> Dict[str, Any]:
        """Extract comparison specific parameters."""
        params = {}
        
        # Look for comparison keywords
        if 'between' in query.lower():
            # Try to extract what's being compared
            between_match = re.search(r'between\s+(\w+)\s+and\s+(\w+)', query, re.IGNORECASE)
            if between_match:
                params['compare_values'] = [between_match.group(1), between_match.group(2)]
        
        # Look for categorical columns for grouping
        categorical_columns = data.select_dtypes(include=['object']).columns.tolist()
        if categorical_columns:
            params['group_by'] = categorical_columns[0]
        
        return params
    
    async def _llm_parameter_extraction(self, query: str, data: pd.DataFrame, 
                                      intent: str) -> Dict[str, Any]:
        """Use LLM to extract complex parameters."""
        try:
            data_info = self._get_data_summary(data)
            
            prompt = f"""
            Extract analysis parameters from this query:
            
            Query: "{query}"
            Intent: {intent}
            
            Data Information:
            {json.dumps(data_info, indent=2)}
            
            Extract relevant parameters such as:
            - Specific columns to analyze
            - Time ranges or periods
            - Thresholds or limits
            - Grouping variables
            - Analysis methods or approaches
            
            Return as JSON with parameter names and values.
            """
            
            response = await self.ollama_client.generate_response(prompt)
            
            try:
                return json.loads(response)
            except:
                return {}
                
        except Exception as e:
            self.logger.error(f"Error in LLM parameter extraction: {e}")
            return {}
    
    async def _generate_clarifications(self, intent: Dict[str, Any], 
                                     parameters: Dict[str, Any], 
                                     data: pd.DataFrame) -> List[str]:
        """Generate clarifying questions if parameters are ambiguous."""
        clarifications = []
        
        primary_intent = intent['primary_intent']
        confidence = intent.get('confidence', 0)
        
        # Low confidence intent detection
        if confidence < 0.5:
            clarifications.append(
                f"I detected your intent as '{primary_intent}' but I'm not very confident. "
                "Could you clarify what type of analysis you're looking for?"
            )
        
        # Missing required parameters
        if primary_intent in self.analysis_mappings:
            mapping = self.analysis_mappings[primary_intent]
            required_params = mapping.get('requires', [])
            
            for param in required_params:
                if param not in parameters:
                    if param == 'time_column':
                        datetime_cols = data.select_dtypes(include=['datetime64']).columns.tolist()
                        if len(datetime_cols) > 1:
                            clarifications.append(
                                f"Which time column should I use for analysis? Options: {', '.join(datetime_cols)}"
                            )
                    elif param == 'value_columns':
                        numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
                        if len(numeric_cols) > 3:
                            clarifications.append(
                                f"Which numeric columns should I analyze? You have: {', '.join(numeric_cols)}"
                            )
        
        # Ambiguous column references
        if 'columns' in parameters:
            referenced_cols = parameters['columns']
            if len(referenced_cols) > 5:
                clarifications.append(
                    f"You referenced many columns ({len(referenced_cols)}). "
                    "Should I focus on specific ones for better analysis?"
                )
        
        return clarifications
    
    def _create_analysis_specification(self, intent: Dict[str, Any], 
                                     parameters: Dict[str, Any], 
                                     data: pd.DataFrame) -> Dict[str, Any]:
        """Create analysis specification from intent and parameters."""
        primary_intent = intent['primary_intent']
        
        if primary_intent not in self.analysis_mappings:
            primary_intent = 'statistical_analysis'  # Default fallback
        
        mapping = self.analysis_mappings[primary_intent]
        
        spec = {
            'agent': mapping['agent'],
            'method': mapping['method'],
            'parameters': {},
            'data_requirements': {
                'columns': parameters.get('columns', []),
                'min_rows': 1,
                'data_types': []
            }
        }
        
        # Map parameters to method parameters
        if primary_intent == 'trend_analysis':
            spec['parameters'] = {
                'time_column': parameters.get('time_column'),
                'value_columns': parameters.get('value_columns', 
                    data.select_dtypes(include=[np.number]).columns.tolist())
            }
        elif primary_intent == 'forecasting':
            spec['parameters'] = {
                'periods': parameters.get('periods', 10)
            }
        elif primary_intent == 'clustering':
            spec['parameters'] = {
                'features': parameters.get('columns', 
                    data.select_dtypes(include=[np.number]).columns.tolist()),
                'n_clusters': parameters.get('n_clusters')
            }
        elif primary_intent == 'pattern_detection':
            spec['parameters'] = {
                'frequency': parameters.get('frequency', 'auto')
            }
        
        return spec
    
    async def _generate_explanation(self, original_query: str, 
                                  analysis_spec: Dict[str, Any]) -> str:
        """Generate explanation of how the query will be processed."""
        try:
            prompt = f"""
            Explain in simple terms how this analysis query will be processed:
            
            Original Query: "{original_query}"
            
            Analysis Plan:
            - Agent: {analysis_spec['agent']}
            - Method: {analysis_spec['method']}
            - Parameters: {json.dumps(analysis_spec['parameters'], indent=2)}
            
            Write a brief, user-friendly explanation of what analysis will be performed.
            """
            
            explanation = await self.ollama_client.generate_response(prompt)
            return explanation
            
        except Exception as e:
            return f"I'll perform {analysis_spec['method']} using the {analysis_spec['agent']} agent."
    
    def _normalize_query(self, query: str) -> str:
        """Normalize query text for processing."""
        # Convert to lowercase
        normalized = query.lower().strip()
        
        # Remove extra whitespace
        normalized = re.sub(r'\s+', ' ', normalized)
        
        # Remove punctuation at the end
        normalized = re.sub(r'[.!?]+$', '', normalized)
        
        return normalized
    
    def _get_data_summary(self, data: pd.DataFrame) -> Dict[str, Any]:
        """Get summary information about the data."""
        return {
            'shape': data.shape,
            'columns': list(data.columns),
            'dtypes': data.dtypes.to_dict(),
            'numeric_columns': data.select_dtypes(include=[np.number]).columns.tolist(),
            'categorical_columns': data.select_dtypes(include=['object']).columns.tolist(),
            'datetime_columns': data.select_dtypes(include=['datetime64']).columns.tolist(),
            'missing_values': data.isnull().sum().to_dict(),
            'sample_values': {col: data[col].dropna().head(3).tolist() 
                            for col in data.columns if not data[col].empty}
        }
    
    def _calculate_query_confidence(self, intent: Dict[str, Any], 
                                  parameters: Dict[str, Any]) -> float:
        """Calculate confidence score for query processing."""
        base_confidence = intent.get('confidence', 0.5)
        
        # Boost confidence if we have clear parameters
        if parameters:
            param_boost = min(0.3, len(parameters) * 0.1)
            base_confidence += param_boost
        
        # Reduce confidence if we need clarifications
        if not parameters:
            base_confidence *= 0.7
        
        return min(1.0, base_confidence)


class QueryResponseFormatter:
    """Formats analysis results into natural language responses."""
    
    def __init__(self, ollama_client: OllamaClient):
        self.ollama_client = ollama_client
        self.logger = logging.getLogger(__name__)
    
    async def format_response(self, query: str, analysis_result: Dict[str, Any]) -> str:
        """Format analysis results into natural language response."""
        try:
            prompt = f"""
            Convert this analysis result into a clear, natural language response for the user:
            
            Original Query: "{query}"
            
            Analysis Results:
            {json.dumps(analysis_result, indent=2, default=str)}
            
            Guidelines:
            1. Start with a direct answer to the user's question
            2. Provide key insights in bullet points
            3. Include relevant numbers and statistics
            4. Explain what the results mean in business terms
            5. Suggest next steps or additional analysis if appropriate
            6. Keep it concise but informative
            7. Use accessible language, avoid technical jargon
            
            Format the response in a friendly, professional tone.
            """
            
            response = await self.ollama_client.generate_response(prompt)
            return response
            
        except Exception as e:
            self.logger.error(f"Error formatting response: {e}")
            return self._create_fallback_response(analysis_result)
    
    def _create_fallback_response(self, analysis_result: Dict[str, Any]) -> str:
        """Create a fallback response when LLM formatting fails."""
        if 'error' in analysis_result:
            return f"I encountered an error during analysis: {analysis_result['error']}"
        
        if 'summary' in analysis_result:
            return f"Analysis completed. {analysis_result['summary']}"
        
        return "Analysis completed successfully. Please check the results in the output sheet."