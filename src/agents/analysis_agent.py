"""
Analysis Agent for Excel-Ollama AI Plugin.
Performs statistical analysis, trend identification, and forecasting.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, r2_score
import json
import asyncio
from datetime import datetime, timedelta

from .base_agent import BaseAgent, AgentMessage


class AnalysisAgent(BaseAgent):
    """Agent responsible for statistical analysis and trend identification."""
    
    def __init__(self, ollama_client):
        super().__init__("analysis", ollama_client)
        self.scaler = StandardScaler()
        
    async def analyze_trends(self, data: pd.DataFrame, time_column: str, 
                           value_columns: List[str]) -> Dict[str, Any]:
        """Analyze trends in time series data."""
        try:
            results = {}
            
            # Ensure time column is datetime
            if time_column in data.columns:
                data[time_column] = pd.to_datetime(data[time_column])
                data = data.sort_values(time_column)
            
            for column in value_columns:
                if column not in data.columns:
                    continue
                    
                # Basic trend analysis
                y = data[column].dropna()
                x = np.arange(len(y))
                
                if len(y) < 2:
                    continue
                
                # Linear regression for trend
                slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
                
                # Trend direction
                trend_direction = "increasing" if slope > 0 else "decreasing" if slope < 0 else "stable"
                
                # Moving averages
                ma_7 = y.rolling(window=min(7, len(y))).mean()
                ma_30 = y.rolling(window=min(30, len(y))).mean()
                
                results[column] = {
                    "slope": slope,
                    "r_squared": r_value**2,
                    "p_value": p_value,
                    "trend_direction": trend_direction,
                    "trend_strength": abs(r_value),
                    "moving_avg_7": ma_7.iloc[-1] if len(ma_7) > 0 else None,
                    "moving_avg_30": ma_30.iloc[-1] if len(ma_30) > 0 else None,
                    "volatility": y.std(),
                    "mean": y.mean(),
                    "median": y.median()
                }
            
            # Generate natural language summary
            summary = await self._generate_trend_summary(results)
            
            return {
                "analysis_type": "trend_analysis",
                "results": results,
                "summary": summary,
                "confidence_score": self._calculate_confidence(results),
                "methodology": "Linear regression with moving averages"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "trend_analysis"}
    
    async def calculate_statistics(self, data: pd.DataFrame) -> Dict[str, Any]:
        """Calculate comprehensive statistics for the dataset."""
        try:
            numeric_columns = data.select_dtypes(include=[np.number]).columns
            results = {}
            
            for column in numeric_columns:
                series = data[column].dropna()
                if len(series) == 0:
                    continue
                
                results[column] = {
                    "count": len(series),
                    "mean": series.mean(),
                    "median": series.median(),
                    "std": series.std(),
                    "min": series.min(),
                    "max": series.max(),
                    "q25": series.quantile(0.25),
                    "q75": series.quantile(0.75),
                    "skewness": stats.skew(series),
                    "kurtosis": stats.kurtosis(series),
                    "coefficient_of_variation": series.std() / series.mean() if series.mean() != 0 else 0
                }
            
            # Correlation matrix for numeric columns
            if len(numeric_columns) > 1:
                correlation_matrix = data[numeric_columns].corr()
                results["correlations"] = correlation_matrix.to_dict()
            
            # Generate summary
            summary = await self._generate_statistics_summary(results)
            
            return {
                "analysis_type": "descriptive_statistics",
                "results": results,
                "summary": summary,
                "confidence_score": 0.95,  # High confidence for descriptive stats
                "methodology": "Descriptive statistics with correlation analysis"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "descriptive_statistics"}
    
    async def perform_forecasting(self, data: pd.DataFrame, periods: int = 10) -> Dict[str, Any]:
        """Perform simple forecasting using linear regression."""
        try:
            numeric_columns = data.select_dtypes(include=[np.number]).columns
            results = {}
            
            for column in numeric_columns:
                series = data[column].dropna()
                if len(series) < 3:  # Need at least 3 points for forecasting
                    continue
                
                # Prepare data
                X = np.arange(len(series)).reshape(-1, 1)
                y = series.values
                
                # Fit model
                model = LinearRegression()
                model.fit(X, y)
                
                # Make predictions
                future_X = np.arange(len(series), len(series) + periods).reshape(-1, 1)
                predictions = model.predict(future_X)
                
                # Calculate confidence intervals (simple approach)
                residuals = y - model.predict(X)
                mse = np.mean(residuals**2)
                std_error = np.sqrt(mse)
                
                results[column] = {
                    "predictions": predictions.tolist(),
                    "confidence_interval_lower": (predictions - 1.96 * std_error).tolist(),
                    "confidence_interval_upper": (predictions + 1.96 * std_error).tolist(),
                    "r_squared": r2_score(y, model.predict(X)),
                    "mse": mse,
                    "trend_slope": model.coef_[0]
                }
            
            # Generate summary
            summary = await self._generate_forecast_summary(results, periods)
            
            return {
                "analysis_type": "forecasting",
                "results": results,
                "summary": summary,
                "confidence_score": self._calculate_forecast_confidence(results),
                "methodology": f"Linear regression forecasting for {periods} periods"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "forecasting"}
    
    async def custom_analysis(self, data: pd.DataFrame, user_query: str) -> Dict[str, Any]:
        """Perform custom analysis based on user query."""
        try:
            # Use Ollama to interpret the query and suggest analysis
            prompt = f"""
            Analyze this data query and suggest appropriate statistical analysis:
            
            Query: {user_query}
            
            Data columns: {list(data.columns)}
            Data shape: {data.shape}
            Data types: {data.dtypes.to_dict()}
            
            Suggest specific analysis methods and provide Python code if needed.
            Focus on practical insights for business users.
            """
            
            response = await self.ollama_client.generate_response(prompt)
            
            # Basic analysis based on common patterns
            results = {}
            
            # If query mentions correlation
            if "correlation" in user_query.lower() or "relationship" in user_query.lower():
                numeric_data = data.select_dtypes(include=[np.number])
                if len(numeric_data.columns) > 1:
                    correlation_matrix = numeric_data.corr()
                    results["correlation_analysis"] = correlation_matrix.to_dict()
            
            # If query mentions trend or time
            if "trend" in user_query.lower() or "time" in user_query.lower():
                date_columns = data.select_dtypes(include=['datetime64']).columns
                numeric_columns = data.select_dtypes(include=[np.number]).columns
                
                if len(date_columns) > 0 and len(numeric_columns) > 0:
                    trend_results = await self.analyze_trends(
                        data, date_columns[0], list(numeric_columns)
                    )
                    results["trend_analysis"] = trend_results
            
            return {
                "analysis_type": "custom_analysis",
                "query": user_query,
                "results": results,
                "llm_response": response,
                "confidence_score": 0.7,  # Lower confidence for custom queries
                "methodology": "LLM-guided analysis with statistical methods"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "custom_analysis"}
    
    async def _generate_trend_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of trend analysis."""
        if not results:
            return "No trend analysis results available."
        
        summaries = []
        for column, data in results.items():
            if isinstance(data, dict) and "trend_direction" in data:
                direction = data["trend_direction"]
                strength = data["trend_strength"]
                r_squared = data["r_squared"]
                
                strength_desc = "strong" if strength > 0.7 else "moderate" if strength > 0.4 else "weak"
                
                summary = f"{column} shows a {strength_desc} {direction} trend (R² = {r_squared:.3f})"
                summaries.append(summary)
        
        return ". ".join(summaries) if summaries else "No clear trends identified."
    
    async def _generate_statistics_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of statistical analysis."""
        if not results:
            return "No statistical analysis results available."
        
        summaries = []
        for column, stats in results.items():
            if isinstance(stats, dict) and "mean" in stats:
                mean = stats["mean"]
                std = stats["std"]
                cv = stats["coefficient_of_variation"]
                
                variability = "high" if cv > 0.5 else "moderate" if cv > 0.2 else "low"
                summary = f"{column}: mean = {mean:.2f}, {variability} variability (CV = {cv:.2f})"
                summaries.append(summary)
        
        return ". ".join(summaries[:3]) if summaries else "No statistical summary available."
    
    async def _generate_forecast_summary(self, results: Dict[str, Any], periods: int) -> str:
        """Generate natural language summary of forecasting results."""
        if not results:
            return f"No forecasting results available for {periods} periods."
        
        summaries = []
        for column, data in results.items():
            if isinstance(data, dict) and "predictions" in data:
                trend_slope = data["trend_slope"]
                r_squared = data["r_squared"]
                
                direction = "increasing" if trend_slope > 0 else "decreasing" if trend_slope < 0 else "stable"
                accuracy = "high" if r_squared > 0.8 else "moderate" if r_squared > 0.5 else "low"
                
                summary = f"{column} forecast shows {direction} trend with {accuracy} accuracy (R² = {r_squared:.3f})"
                summaries.append(summary)
        
        return ". ".join(summaries) if summaries else f"Forecast generated for {periods} periods."
    
    def _calculate_confidence(self, results: Dict[str, Any]) -> float:
        """Calculate overall confidence score for analysis results."""
        if not results:
            return 0.0
        
        confidence_scores = []
        for column, data in results.items():
            if isinstance(data, dict):
                if "r_squared" in data:
                    confidence_scores.append(data["r_squared"])
                elif "p_value" in data:
                    # Convert p-value to confidence (1 - p_value)
                    confidence_scores.append(max(0, 1 - data["p_value"]))
        
        return np.mean(confidence_scores) if confidence_scores else 0.5
    
    def _calculate_forecast_confidence(self, results: Dict[str, Any]) -> float:
        """Calculate confidence score for forecasting results."""
        if not results:
            return 0.0
        
        r_squared_values = []
        for column, data in results.items():
            if isinstance(data, dict) and "r_squared" in data:
                r_squared_values.append(data["r_squared"])
        
        return np.mean(r_squared_values) if r_squared_values else 0.3