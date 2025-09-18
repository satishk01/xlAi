"""
Pattern Agent for Excel-Ollama AI Plugin.
Detects patterns, anomalies, and clusters in activity data.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple
from sklearn.cluster import KMeans, DBSCAN
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from scipy import stats
from scipy.signal import find_peaks
import asyncio
from datetime import datetime

from .base_agent import BaseAgent, AgentMessage


class PatternAgent(BaseAgent):
    """Agent responsible for pattern detection and anomaly identification."""
    
    def __init__(self, ollama_client):
        super().__init__("pattern", ollama_client)
        self.scaler = StandardScaler()
        
    async def detect_seasonal_patterns(self, data: pd.DataFrame, 
                                     frequency: str = 'auto') -> Dict[str, Any]:
        """Detect seasonal patterns in time series data."""
        try:
            results = {}
            
            # Find datetime columns
            datetime_cols = data.select_dtypes(include=['datetime64']).columns
            numeric_cols = data.select_dtypes(include=[np.number]).columns
            
            if len(datetime_cols) == 0 or len(numeric_cols) == 0:
                return {"error": "Need datetime and numeric columns for seasonal analysis"}
            
            time_col = datetime_cols[0]
            data_sorted = data.sort_values(time_col)
            
            for col in numeric_cols:
                series = data_sorted[col].dropna()
                if len(series) < 10:  # Need sufficient data points
                    continue
                
                # Extract time components
                dates = data_sorted[time_col].iloc[:len(series)]
                
                # Detect patterns by time components
                patterns = {}
                
                # Daily patterns (hour of day)
                if len(series) > 24:
                    hourly_avg = series.groupby(dates.dt.hour).mean()
                    hourly_std = series.groupby(dates.dt.hour).std()
                    patterns['hourly'] = {
                        'averages': hourly_avg.to_dict(),
                        'peak_hours': hourly_avg.nlargest(3).index.tolist(),
                        'low_hours': hourly_avg.nsmallest(3).index.tolist()
                    }
                
                # Weekly patterns (day of week)
                if len(series) > 7:
                    daily_avg = series.groupby(dates.dt.dayofweek).mean()
                    patterns['weekly'] = {
                        'averages': daily_avg.to_dict(),
                        'peak_days': daily_avg.nlargest(2).index.tolist(),
                        'low_days': daily_avg.nsmallest(2).index.tolist()
                    }
                
                # Monthly patterns
                if len(series) > 30:
                    monthly_avg = series.groupby(dates.dt.month).mean()
                    patterns['monthly'] = {
                        'averages': monthly_avg.to_dict(),
                        'peak_months': monthly_avg.nlargest(3).index.tolist(),
                        'low_months': monthly_avg.nsmallest(3).index.tolist()
                    }
                
                # Detect cyclical patterns using autocorrelation
                if len(series) > 50:
                    autocorr = self._calculate_autocorrelation(series.values)
                    patterns['cycles'] = self._find_cycles(autocorr)
                
                results[col] = patterns
            
            # Generate summary
            summary = await self._generate_seasonal_summary(results)
            
            return {
                "analysis_type": "seasonal_patterns",
                "results": results,
                "summary": summary,
                "confidence_score": self._calculate_pattern_confidence(results),
                "methodology": "Time-based grouping with autocorrelation analysis"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "seasonal_patterns"}
    
    async def identify_outliers(self, data: pd.DataFrame, 
                              method: str = 'isolation_forest') -> Dict[str, Any]:
        """Identify outliers and anomalies in the data."""
        try:
            numeric_data = data.select_dtypes(include=[np.number])
            if numeric_data.empty:
                return {"error": "No numeric data available for outlier detection"}
            
            results = {}
            
            for column in numeric_data.columns:
                series = numeric_data[column].dropna()
                if len(series) < 10:
                    continue
                
                outliers = {}
                
                # Statistical outliers (Z-score method)
                z_scores = np.abs(stats.zscore(series))
                z_outliers = np.where(z_scores > 3)[0]
                outliers['z_score'] = {
                    'indices': z_outliers.tolist(),
                    'values': series.iloc[z_outliers].tolist(),
                    'z_scores': z_scores[z_outliers].tolist()
                }
                
                # IQR method
                Q1 = series.quantile(0.25)
                Q3 = series.quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                iqr_outliers = series[(series < lower_bound) | (series > upper_bound)]
                outliers['iqr'] = {
                    'indices': iqr_outliers.index.tolist(),
                    'values': iqr_outliers.tolist(),
                    'bounds': {'lower': lower_bound, 'upper': upper_bound}
                }
                
                # Isolation Forest (if enough data)
                if len(series) >= 20 and method == 'isolation_forest':
                    iso_forest = IsolationForest(contamination=0.1, random_state=42)
                    outlier_labels = iso_forest.fit_predict(series.values.reshape(-1, 1))
                    iso_outliers = np.where(outlier_labels == -1)[0]
                    
                    outliers['isolation_forest'] = {
                        'indices': iso_outliers.tolist(),
                        'values': series.iloc[iso_outliers].tolist(),
                        'scores': iso_forest.decision_function(series.values.reshape(-1, 1))[iso_outliers].tolist()
                    }
                
                results[column] = outliers
            
            # Generate summary
            summary = await self._generate_outlier_summary(results)
            
            return {
                "analysis_type": "outlier_detection",
                "results": results,
                "summary": summary,
                "confidence_score": 0.8,
                "methodology": f"Multiple methods: Z-score, IQR, {method}"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "outlier_detection"}
    
    async def cluster_activities(self, data: pd.DataFrame, 
                               features: List[str] = None) -> Dict[str, Any]:
        """Cluster similar activities or data points."""
        try:
            if features is None:
                features = data.select_dtypes(include=[np.number]).columns.tolist()
            
            # Filter data to specified features
            cluster_data = data[features].dropna()
            
            if cluster_data.empty or len(cluster_data) < 3:
                return {"error": "Insufficient data for clustering"}
            
            # Standardize the data
            scaled_data = self.scaler.fit_transform(cluster_data)
            
            results = {}
            
            # K-means clustering (try different k values)
            kmeans_results = {}
            inertias = []
            silhouette_scores = []
            
            max_k = min(10, len(cluster_data) // 2)
            
            for k in range(2, max_k + 1):
                kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
                cluster_labels = kmeans.fit_predict(scaled_data)
                
                inertias.append(kmeans.inertia_)
                
                # Calculate silhouette score
                if len(np.unique(cluster_labels)) > 1:
                    from sklearn.metrics import silhouette_score
                    sil_score = silhouette_score(scaled_data, cluster_labels)
                    silhouette_scores.append(sil_score)
                else:
                    silhouette_scores.append(0)
                
                kmeans_results[k] = {
                    'labels': cluster_labels.tolist(),
                    'centers': kmeans.cluster_centers_.tolist(),
                    'inertia': kmeans.inertia_,
                    'silhouette_score': silhouette_scores[-1]
                }
            
            # Find optimal k using elbow method
            optimal_k = self._find_optimal_k(inertias, silhouette_scores)
            
            # DBSCAN clustering
            dbscan = DBSCAN(eps=0.5, min_samples=3)
            dbscan_labels = dbscan.fit_predict(scaled_data)
            n_clusters_dbscan = len(set(dbscan_labels)) - (1 if -1 in dbscan_labels else 0)
            
            results = {
                'kmeans': kmeans_results,
                'optimal_k': optimal_k,
                'dbscan': {
                    'labels': dbscan_labels.tolist(),
                    'n_clusters': n_clusters_dbscan,
                    'n_noise': list(dbscan_labels).count(-1)
                },
                'features_used': features,
                'data_shape': cluster_data.shape
            }
            
            # Generate cluster profiles
            if optimal_k in kmeans_results:
                cluster_profiles = self._generate_cluster_profiles(
                    cluster_data, kmeans_results[optimal_k]['labels'], features
                )
                results['cluster_profiles'] = cluster_profiles
            
            # Generate summary
            summary = await self._generate_clustering_summary(results)
            
            return {
                "analysis_type": "clustering",
                "results": results,
                "summary": summary,
                "confidence_score": max(silhouette_scores) if silhouette_scores else 0.5,
                "methodology": "K-means and DBSCAN clustering with optimal k selection"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "clustering"}
    
    async def analyze_behavioral_patterns(self, data: pd.DataFrame) -> Dict[str, Any]:
        """Analyze behavioral patterns in activity sequences."""
        try:
            results = {}
            
            # Look for sequence patterns if there's a time component
            datetime_cols = data.select_dtypes(include=['datetime64']).columns
            
            if len(datetime_cols) > 0:
                time_col = datetime_cols[0]
                data_sorted = data.sort_values(time_col)
                
                # Analyze activity frequency patterns
                if 'activity' in data.columns or 'event' in data.columns:
                    activity_col = 'activity' if 'activity' in data.columns else 'event'
                    
                    # Frequency analysis
                    activity_counts = data_sorted[activity_col].value_counts()
                    results['activity_frequency'] = activity_counts.to_dict()
                    
                    # Time-based patterns
                    data_sorted['hour'] = data_sorted[time_col].dt.hour
                    data_sorted['day_of_week'] = data_sorted[time_col].dt.dayofweek
                    
                    hourly_patterns = data_sorted.groupby(['hour', activity_col]).size().unstack(fill_value=0)
                    daily_patterns = data_sorted.groupby(['day_of_week', activity_col]).size().unstack(fill_value=0)
                    
                    results['hourly_patterns'] = hourly_patterns.to_dict()
                    results['daily_patterns'] = daily_patterns.to_dict()
                
                # Analyze intervals between events
                time_diffs = data_sorted[time_col].diff().dt.total_seconds() / 3600  # Convert to hours
                time_diffs = time_diffs.dropna()
                
                if len(time_diffs) > 0:
                    results['time_intervals'] = {
                        'mean_hours': time_diffs.mean(),
                        'median_hours': time_diffs.median(),
                        'std_hours': time_diffs.std(),
                        'min_hours': time_diffs.min(),
                        'max_hours': time_diffs.max()
                    }
            
            # Analyze numeric patterns
            numeric_cols = data.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                series = data[col].dropna()
                if len(series) < 5:
                    continue
                
                # Find peaks and valleys
                peaks, _ = find_peaks(series.values)
                valleys, _ = find_peaks(-series.values)
                
                results[f'{col}_patterns'] = {
                    'peaks': {'indices': peaks.tolist(), 'values': series.iloc[peaks].tolist()},
                    'valleys': {'indices': valleys.tolist(), 'values': series.iloc[valleys].tolist()},
                    'peak_frequency': len(peaks) / len(series) if len(series) > 0 else 0,
                    'average_peak_value': series.iloc[peaks].mean() if len(peaks) > 0 else None,
                    'average_valley_value': series.iloc[valleys].mean() if len(valleys) > 0 else None
                }
            
            # Generate summary
            summary = await self._generate_behavioral_summary(results)
            
            return {
                "analysis_type": "behavioral_patterns",
                "results": results,
                "summary": summary,
                "confidence_score": 0.7,
                "methodology": "Time-based pattern analysis with peak detection"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "behavioral_patterns"}
    
    def _calculate_autocorrelation(self, series: np.ndarray, max_lags: int = None) -> np.ndarray:
        """Calculate autocorrelation function."""
        if max_lags is None:
            max_lags = min(len(series) // 4, 50)
        
        autocorr = np.correlate(series, series, mode='full')
        autocorr = autocorr[autocorr.size // 2:]
        autocorr = autocorr / autocorr[0]  # Normalize
        
        return autocorr[:max_lags]
    
    def _find_cycles(self, autocorr: np.ndarray) -> Dict[str, Any]:
        """Find cyclical patterns in autocorrelation."""
        peaks, properties = find_peaks(autocorr[1:], height=0.3, distance=3)
        peaks += 1  # Adjust for skipping first element
        
        cycles = []
        for peak in peaks:
            cycles.append({
                'period': int(peak),
                'strength': float(autocorr[peak])
            })
        
        return {
            'detected_cycles': cycles,
            'strongest_cycle': max(cycles, key=lambda x: x['strength']) if cycles else None
        }
    
    def _find_optimal_k(self, inertias: List[float], silhouette_scores: List[float]) -> int:
        """Find optimal number of clusters using elbow method and silhouette score."""
        if not silhouette_scores:
            return 2
        
        # Find k with highest silhouette score
        best_silhouette_k = silhouette_scores.index(max(silhouette_scores)) + 2
        
        # Simple elbow method - find point of maximum curvature
        if len(inertias) >= 3:
            diffs = np.diff(inertias)
            diff2 = np.diff(diffs)
            elbow_k = np.argmax(diff2) + 3  # +3 because we start from k=2 and take second derivative
            
            # Choose between silhouette and elbow method
            if abs(best_silhouette_k - elbow_k) <= 2:
                return best_silhouette_k
            else:
                return best_silhouette_k if max(silhouette_scores) > 0.5 else elbow_k
        
        return best_silhouette_k
    
    def _generate_cluster_profiles(self, data: pd.DataFrame, labels: List[int], 
                                 features: List[str]) -> Dict[str, Any]:
        """Generate profiles for each cluster."""
        profiles = {}
        
        for cluster_id in set(labels):
            cluster_mask = np.array(labels) == cluster_id
            cluster_data = data[cluster_mask]
            
            profile = {}
            for feature in features:
                if feature in cluster_data.columns:
                    profile[feature] = {
                        'mean': cluster_data[feature].mean(),
                        'std': cluster_data[feature].std(),
                        'min': cluster_data[feature].min(),
                        'max': cluster_data[feature].max(),
                        'count': len(cluster_data)
                    }
            
            profiles[f'cluster_{cluster_id}'] = profile
        
        return profiles
    
    async def _generate_seasonal_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of seasonal patterns."""
        if not results:
            return "No seasonal patterns detected."
        
        summaries = []
        for column, patterns in results.items():
            if isinstance(patterns, dict):
                pattern_types = []
                
                if 'hourly' in patterns and patterns['hourly'].get('peak_hours'):
                    peak_hours = patterns['hourly']['peak_hours']
                    pattern_types.append(f"peaks at hours {peak_hours}")
                
                if 'weekly' in patterns and patterns['weekly'].get('peak_days'):
                    peak_days = patterns['weekly']['peak_days']
                    day_names = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
                    peak_day_names = [day_names[d] for d in peak_days if d < 7]
                    pattern_types.append(f"weekly peaks on {', '.join(peak_day_names)}")
                
                if 'cycles' in patterns and patterns['cycles'].get('strongest_cycle'):
                    cycle = patterns['cycles']['strongest_cycle']
                    pattern_types.append(f"cyclical pattern every {cycle['period']} periods")
                
                if pattern_types:
                    summary = f"{column}: {', '.join(pattern_types)}"
                    summaries.append(summary)
        
        return ". ".join(summaries) if summaries else "No clear seasonal patterns found."
    
    async def _generate_outlier_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of outlier detection."""
        if not results:
            return "No outliers detected."
        
        summaries = []
        for column, outliers in results.items():
            if isinstance(outliers, dict):
                total_outliers = set()
                
                for method, data in outliers.items():
                    if isinstance(data, dict) and 'indices' in data:
                        total_outliers.update(data['indices'])
                
                if total_outliers:
                    summary = f"{column}: {len(total_outliers)} outliers detected"
                    summaries.append(summary)
        
        return ". ".join(summaries) if summaries else "No significant outliers found."
    
    async def _generate_clustering_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of clustering results."""
        if not results:
            return "No clustering results available."
        
        optimal_k = results.get('optimal_k', 0)
        dbscan_clusters = results.get('dbscan', {}).get('n_clusters', 0)
        
        summary_parts = []
        
        if optimal_k > 0:
            summary_parts.append(f"K-means identified {optimal_k} optimal clusters")
        
        if dbscan_clusters > 0:
            summary_parts.append(f"DBSCAN found {dbscan_clusters} density-based clusters")
        
        return ". ".join(summary_parts) if summary_parts else "No clear clustering structure found."
    
    async def _generate_behavioral_summary(self, results: Dict[str, Any]) -> str:
        """Generate natural language summary of behavioral patterns."""
        if not results:
            return "No behavioral patterns detected."
        
        summaries = []
        
        if 'activity_frequency' in results:
            top_activity = max(results['activity_frequency'].items(), key=lambda x: x[1])
            summaries.append(f"Most frequent activity: {top_activity[0]} ({top_activity[1]} occurrences)")
        
        if 'time_intervals' in results:
            mean_hours = results['time_intervals']['mean_hours']
            summaries.append(f"Average time between events: {mean_hours:.1f} hours")
        
        # Look for peak patterns
        for key, value in results.items():
            if key.endswith('_patterns') and isinstance(value, dict):
                if 'peaks' in value and value['peaks']['values']:
                    avg_peak = np.mean(value['peaks']['values'])
                    summaries.append(f"{key.replace('_patterns', '')} shows peaks averaging {avg_peak:.2f}")
        
        return ". ".join(summaries[:3]) if summaries else "No clear behavioral patterns identified."
    
    def _calculate_pattern_confidence(self, results: Dict[str, Any]) -> float:
        """Calculate confidence score for pattern detection."""
        if not results:
            return 0.0
        
        confidence_scores = []
        
        for column, patterns in results.items():
            if isinstance(patterns, dict):
                # Count detected pattern types
                pattern_count = 0
                if 'hourly' in patterns and patterns['hourly'].get('peak_hours'):
                    pattern_count += 1
                if 'weekly' in patterns and patterns['weekly'].get('peak_days'):
                    pattern_count += 1
                if 'monthly' in patterns and patterns['monthly'].get('peak_months'):
                    pattern_count += 1
                if 'cycles' in patterns and patterns['cycles'].get('detected_cycles'):
                    pattern_count += len(patterns['cycles']['detected_cycles'])
                
                # More patterns = higher confidence
                confidence_scores.append(min(1.0, pattern_count * 0.25))
        
        return np.mean(confidence_scores) if confidence_scores else 0.3