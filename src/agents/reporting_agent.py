"""
Reporting Agent for Excel-Ollama AI Plugin.
Generates insights, reports, and visualization recommendations.
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional, Tuple
import json
import asyncio
from datetime import datetime

from .base_agent import BaseAgent, AgentMessage


class ReportingAgent(BaseAgent):
    """Agent responsible for generating reports and insights."""
    
    def __init__(self, ollama_client):
        super().__init__("reporting", ollama_client)
        self.report_templates = self._load_report_templates()
        
    async def generate_summary(self, analysis_results: Dict[str, Any]) -> Dict[str, Any]:
        """Generate natural language summary from analysis results."""
        try:
            # Combine results from different agents
            combined_insights = []
            key_metrics = {}
            
            for agent_type, results in analysis_results.items():
                if isinstance(results, dict) and 'summary' in results:
                    combined_insights.append(f"{agent_type.title()}: {results['summary']}")
                    
                    # Extract key metrics
                    if 'results' in results:
                        key_metrics[agent_type] = self._extract_key_metrics(results['results'])
            
            # Generate comprehensive summary using Ollama
            prompt = f"""
            Create a comprehensive business summary from these analysis results:
            
            Analysis Insights:
            {chr(10).join(combined_insights)}
            
            Key Metrics:
            {json.dumps(key_metrics, indent=2)}
            
            Please provide:
            1. Executive summary (2-3 sentences)
            2. Key findings (3-5 bullet points)
            3. Actionable recommendations (2-3 suggestions)
            4. Areas for further investigation
            
            Write in clear, business-friendly language suitable for executives.
            """
            
            llm_summary = await self.ollama_client.generate_response(prompt)
            
            return {
                "analysis_type": "comprehensive_summary",
                "executive_summary": self._extract_executive_summary(llm_summary),
                "key_findings": self._extract_key_findings(llm_summary),
                "recommendations": self._extract_recommendations(llm_summary),
                "detailed_insights": combined_insights,
                "key_metrics": key_metrics,
                "llm_response": llm_summary,
                "confidence_score": self._calculate_summary_confidence(analysis_results),
                "methodology": "Multi-agent analysis synthesis with LLM interpretation"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "comprehensive_summary"}
    
    async def create_report(self, template: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Create formatted report using specified template."""
        try:
            if template not in self.report_templates:
                return {"error": f"Template '{template}' not found"}
            
            template_config = self.report_templates[template]
            
            # Generate report sections
            report_sections = {}
            
            for section in template_config['sections']:
                section_content = await self._generate_section_content(section, data)
                report_sections[section['name']] = section_content
            
            # Format the complete report
            formatted_report = self._format_report(template_config, report_sections)
            
            return {
                "analysis_type": "formatted_report",
                "template_used": template,
                "report_content": formatted_report,
                "sections": report_sections,
                "metadata": {
                    "generated_at": datetime.now().isoformat(),
                    "data_summary": self._summarize_input_data(data)
                },
                "confidence_score": 0.85,
                "methodology": f"Template-based report generation using {template} format"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "formatted_report"}
    
    async def recommend_visualizations(self, data_characteristics: Dict[str, Any]) -> Dict[str, Any]:
        """Recommend appropriate visualizations based on data characteristics."""
        try:
            recommendations = []
            
            # Analyze data characteristics
            data_types = data_characteristics.get('data_types', {})
            data_shape = data_characteristics.get('shape', (0, 0))
            has_time_series = data_characteristics.get('has_time_series', False)
            numeric_columns = data_characteristics.get('numeric_columns', [])
            categorical_columns = data_characteristics.get('categorical_columns', [])
            
            # Time series visualizations
            if has_time_series and len(numeric_columns) > 0:
                recommendations.append({
                    'type': 'line_chart',
                    'title': 'Time Series Trend Analysis',
                    'description': 'Shows trends over time for numeric variables',
                    'columns': numeric_columns[:3],  # Limit to 3 series
                    'priority': 'high',
                    'chart_config': {
                        'x_axis': 'time',
                        'y_axis': numeric_columns[0] if numeric_columns else 'value',
                        'show_trend_line': True
                    }
                })
            
            # Distribution visualizations
            if len(numeric_columns) > 0:
                recommendations.append({
                    'type': 'histogram',
                    'title': 'Data Distribution Analysis',
                    'description': 'Shows distribution of numeric variables',
                    'columns': numeric_columns[:2],
                    'priority': 'medium',
                    'chart_config': {
                        'bins': 20,
                        'show_normal_curve': True
                    }
                })
                
                # Box plot for outlier detection
                recommendations.append({
                    'type': 'box_plot',
                    'title': 'Outlier Detection',
                    'description': 'Identifies outliers and quartile ranges',
                    'columns': numeric_columns[:4],
                    'priority': 'medium',
                    'chart_config': {
                        'show_outliers': True,
                        'show_mean': True
                    }
                })
            
            # Correlation analysis
            if len(numeric_columns) >= 2:
                recommendations.append({
                    'type': 'correlation_heatmap',
                    'title': 'Correlation Matrix',
                    'description': 'Shows relationships between numeric variables',
                    'columns': numeric_columns,
                    'priority': 'high',
                    'chart_config': {
                        'color_scheme': 'RdBu',
                        'show_values': True
                    }
                })
                
                # Scatter plot for top correlations
                recommendations.append({
                    'type': 'scatter_plot',
                    'title': 'Variable Relationships',
                    'description': 'Scatter plot of highly correlated variables',
                    'columns': numeric_columns[:2],
                    'priority': 'medium',
                    'chart_config': {
                        'show_regression_line': True,
                        'show_confidence_interval': True
                    }
                })
            
            # Categorical data visualizations
            if len(categorical_columns) > 0:
                recommendations.append({
                    'type': 'bar_chart',
                    'title': 'Category Distribution',
                    'description': 'Shows frequency of categorical variables',
                    'columns': categorical_columns[:1],
                    'priority': 'medium',
                    'chart_config': {
                        'sort_by': 'frequency',
                        'show_percentages': True
                    }
                })
                
                # Pie chart for single categorical variable
                if len(categorical_columns) >= 1:
                    recommendations.append({
                        'type': 'pie_chart',
                        'title': 'Category Proportions',
                        'description': 'Shows proportional breakdown of categories',
                        'columns': [categorical_columns[0]],
                        'priority': 'low',
                        'chart_config': {
                            'show_percentages': True,
                            'max_categories': 8
                        }
                    })
            
            # Mixed data visualizations
            if len(numeric_columns) > 0 and len(categorical_columns) > 0:
                recommendations.append({
                    'type': 'grouped_bar_chart',
                    'title': 'Category Comparison',
                    'description': 'Compares numeric values across categories',
                    'columns': [categorical_columns[0], numeric_columns[0]],
                    'priority': 'high',
                    'chart_config': {
                        'group_by': categorical_columns[0],
                        'aggregate': 'mean'
                    }
                })
            
            # Sort recommendations by priority
            priority_order = {'high': 3, 'medium': 2, 'low': 1}
            recommendations.sort(key=lambda x: priority_order.get(x['priority'], 0), reverse=True)
            
            # Generate LLM-based recommendations
            llm_recommendations = await self._generate_llm_visualization_recommendations(data_characteristics)
            
            return {
                "analysis_type": "visualization_recommendations",
                "recommendations": recommendations,
                "llm_suggestions": llm_recommendations,
                "data_summary": data_characteristics,
                "confidence_score": 0.8,
                "methodology": "Rule-based recommendations with LLM enhancement"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "visualization_recommendations"}
    
    async def build_dashboard(self, key_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Build executive dashboard with key performance indicators."""
        try:
            dashboard_components = []
            
            # KPI Cards
            kpi_cards = self._create_kpi_cards(key_metrics)
            if kpi_cards:
                dashboard_components.append({
                    'type': 'kpi_cards',
                    'title': 'Key Performance Indicators',
                    'components': kpi_cards,
                    'layout': {'columns': min(4, len(kpi_cards)), 'height': 150}
                })
            
            # Trend Charts
            trend_charts = self._create_trend_charts(key_metrics)
            if trend_charts:
                dashboard_components.append({
                    'type': 'trend_charts',
                    'title': 'Performance Trends',
                    'components': trend_charts,
                    'layout': {'columns': 2, 'height': 300}
                })
            
            # Distribution Charts
            distribution_charts = self._create_distribution_charts(key_metrics)
            if distribution_charts:
                dashboard_components.append({
                    'type': 'distribution_charts',
                    'title': 'Data Distribution',
                    'components': distribution_charts,
                    'layout': {'columns': 2, 'height': 250}
                })
            
            # Summary Table
            summary_table = self._create_summary_table(key_metrics)
            if summary_table:
                dashboard_components.append({
                    'type': 'summary_table',
                    'title': 'Detailed Metrics',
                    'components': [summary_table],
                    'layout': {'columns': 1, 'height': 200}
                })
            
            # Generate dashboard insights
            insights = await self._generate_dashboard_insights(key_metrics)
            
            return {
                "analysis_type": "executive_dashboard",
                "dashboard_components": dashboard_components,
                "insights": insights,
                "metadata": {
                    "created_at": datetime.now().isoformat(),
                    "metrics_count": len(key_metrics),
                    "component_count": len(dashboard_components)
                },
                "confidence_score": 0.9,
                "methodology": "Automated dashboard generation with KPI identification"
            }
            
        except Exception as e:
            return {"error": str(e), "analysis_type": "executive_dashboard"}
    
    def _load_report_templates(self) -> Dict[str, Any]:
        """Load predefined report templates."""
        return {
            'executive_summary': {
                'name': 'Executive Summary Report',
                'sections': [
                    {'name': 'overview', 'type': 'text', 'required': True},
                    {'name': 'key_metrics', 'type': 'metrics', 'required': True},
                    {'name': 'trends', 'type': 'analysis', 'required': False},
                    {'name': 'recommendations', 'type': 'text', 'required': True}
                ]
            },
            'detailed_analysis': {
                'name': 'Detailed Analysis Report',
                'sections': [
                    {'name': 'data_summary', 'type': 'data', 'required': True},
                    {'name': 'statistical_analysis', 'type': 'analysis', 'required': True},
                    {'name': 'pattern_analysis', 'type': 'analysis', 'required': False},
                    {'name': 'anomalies', 'type': 'analysis', 'required': False},
                    {'name': 'conclusions', 'type': 'text', 'required': True}
                ]
            },
            'performance_report': {
                'name': 'Performance Report',
                'sections': [
                    {'name': 'kpis', 'type': 'metrics', 'required': True},
                    {'name': 'trends', 'type': 'analysis', 'required': True},
                    {'name': 'comparisons', 'type': 'analysis', 'required': False},
                    {'name': 'action_items', 'type': 'text', 'required': True}
                ]
            }
        }
    
    async def _generate_section_content(self, section: Dict[str, Any], data: Dict[str, Any]) -> str:
        """Generate content for a specific report section."""
        section_type = section['type']
        section_name = section['name']
        
        if section_type == 'text':
            # Generate text content using LLM
            prompt = f"""
            Generate a {section_name} section for a business report based on this data:
            {json.dumps(data, indent=2, default=str)}
            
            Make it concise, professional, and actionable.
            """
            return await self.ollama_client.generate_response(prompt)
        
        elif section_type == 'metrics':
            # Format key metrics
            metrics = self._extract_key_metrics(data)
            return self._format_metrics_section(metrics)
        
        elif section_type == 'analysis':
            # Extract analysis results
            analysis_content = []
            for key, value in data.items():
                if isinstance(value, dict) and 'summary' in value:
                    analysis_content.append(f"**{key.title()}**: {value['summary']}")
            return '\n\n'.join(analysis_content)
        
        elif section_type == 'data':
            # Summarize data characteristics
            return self._format_data_summary(data)
        
        return f"Content for {section_name} section"
    
    def _format_report(self, template_config: Dict[str, Any], sections: Dict[str, str]) -> str:
        """Format complete report from sections."""
        report_lines = [
            f"# {template_config['name']}",
            f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "---",
            ""
        ]
        
        for section in template_config['sections']:
            section_name = section['name']
            if section_name in sections:
                report_lines.extend([
                    f"## {section_name.replace('_', ' ').title()}",
                    "",
                    sections[section_name],
                    "",
                    "---",
                    ""
                ])
        
        return '\n'.join(report_lines)
    
    def _extract_key_metrics(self, data: Any) -> Dict[str, Any]:
        """Extract key metrics from analysis results."""
        metrics = {}
        
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, (int, float)):
                    metrics[key] = value
                elif isinstance(value, dict):
                    # Look for common metric patterns
                    if 'mean' in value:
                        metrics[f"{key}_mean"] = value['mean']
                    if 'count' in value:
                        metrics[f"{key}_count"] = value['count']
                    if 'confidence_score' in value:
                        metrics[f"{key}_confidence"] = value['confidence_score']
        
        return metrics
    
    def _format_metrics_section(self, metrics: Dict[str, Any]) -> str:
        """Format metrics into a readable section."""
        if not metrics:
            return "No key metrics available."
        
        lines = ["### Key Metrics", ""]
        
        for metric, value in metrics.items():
            formatted_name = metric.replace('_', ' ').title()
            if isinstance(value, float):
                formatted_value = f"{value:.3f}"
            else:
                formatted_value = str(value)
            
            lines.append(f"- **{formatted_name}**: {formatted_value}")
        
        return '\n'.join(lines)
    
    def _format_data_summary(self, data: Dict[str, Any]) -> str:
        """Format data summary section."""
        lines = ["### Data Summary", ""]
        
        # Look for data characteristics
        if 'shape' in data:
            lines.append(f"- **Dataset Size**: {data['shape'][0]} rows, {data['shape'][1]} columns")
        
        if 'data_types' in data:
            lines.append(f"- **Data Types**: {len(data['data_types'])} different types")
        
        if 'missing_values' in data:
            lines.append(f"- **Missing Values**: {data['missing_values']} total")
        
        return '\n'.join(lines)
    
    def _create_kpi_cards(self, key_metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create KPI cards for dashboard."""
        kpi_cards = []
        
        # Look for important metrics
        for key, value in key_metrics.items():
            if isinstance(value, (int, float)):
                card = {
                    'title': key.replace('_', ' ').title(),
                    'value': value,
                    'format': 'number' if isinstance(value, int) else 'decimal',
                    'trend': self._determine_trend(key, value),
                    'color': self._determine_color(key, value)
                }
                kpi_cards.append(card)
        
        return kpi_cards[:6]  # Limit to 6 KPIs
    
    def _create_trend_charts(self, key_metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create trend charts for dashboard."""
        # This would typically use historical data
        # For now, return placeholder structure
        return [
            {
                'title': 'Performance Trend',
                'type': 'line',
                'data': 'placeholder_trend_data',
                'config': {'show_trend_line': True}
            }
        ]
    
    def _create_distribution_charts(self, key_metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Create distribution charts for dashboard."""
        return [
            {
                'title': 'Metric Distribution',
                'type': 'histogram',
                'data': 'placeholder_distribution_data',
                'config': {'bins': 20}
            }
        ]
    
    def _create_summary_table(self, key_metrics: Dict[str, Any]) -> Dict[str, Any]:
        """Create summary table for dashboard."""
        return {
            'title': 'Metrics Summary',
            'type': 'table',
            'data': key_metrics,
            'config': {'sortable': True, 'searchable': True}
        }
    
    async def _generate_dashboard_insights(self, key_metrics: Dict[str, Any]) -> str:
        """Generate insights for dashboard."""
        prompt = f"""
        Generate 3-5 key insights from these dashboard metrics:
        {json.dumps(key_metrics, indent=2, default=str)}
        
        Focus on:
        1. Notable patterns or trends
        2. Areas of concern or opportunity
        3. Actionable recommendations
        
        Keep insights concise and business-focused.
        """
        
        return await self.ollama_client.generate_response(prompt)
    
    def _determine_trend(self, metric_name: str, value: float) -> str:
        """Determine trend direction for KPI (placeholder logic)."""
        # This would typically compare with historical values
        return 'stable'
    
    def _determine_color(self, metric_name: str, value: float) -> str:
        """Determine color for KPI based on value."""
        if 'error' in metric_name.lower() or 'fail' in metric_name.lower():
            return 'red' if value > 0 else 'green'
        elif 'success' in metric_name.lower() or 'complete' in metric_name.lower():
            return 'green' if value > 0.8 else 'yellow' if value > 0.5 else 'red'
        else:
            return 'blue'  # Default color
    
    async def _generate_llm_visualization_recommendations(self, data_characteristics: Dict[str, Any]) -> str:
        """Generate LLM-based visualization recommendations."""
        prompt = f"""
        Based on these data characteristics, recommend the best visualizations:
        
        Data Characteristics:
        {json.dumps(data_characteristics, indent=2)}
        
        Provide specific visualization recommendations with reasoning.
        Consider the audience (business users) and the goal (insights discovery).
        """
        
        return await self.ollama_client.generate_response(prompt)
    
    def _extract_executive_summary(self, llm_response: str) -> str:
        """Extract executive summary from LLM response."""
        lines = llm_response.split('\n')
        for i, line in enumerate(lines):
            if 'executive summary' in line.lower():
                # Return next few lines
                return '\n'.join(lines[i+1:i+4]).strip()
        return llm_response[:200] + "..."  # Fallback to first 200 chars
    
    def _extract_key_findings(self, llm_response: str) -> List[str]:
        """Extract key findings from LLM response."""
        findings = []
        lines = llm_response.split('\n')
        
        in_findings_section = False
        for line in lines:
            if 'key findings' in line.lower() or 'findings' in line.lower():
                in_findings_section = True
                continue
            elif in_findings_section and line.strip().startswith(('•', '-', '*', '1.', '2.', '3.')):
                findings.append(line.strip())
            elif in_findings_section and line.strip() == '':
                continue
            elif in_findings_section and not line.strip().startswith(('•', '-', '*')) and len(findings) > 0:
                break
        
        return findings[:5]  # Limit to 5 findings
    
    def _extract_recommendations(self, llm_response: str) -> List[str]:
        """Extract recommendations from LLM response."""
        recommendations = []
        lines = llm_response.split('\n')
        
        in_recommendations_section = False
        for line in lines:
            if 'recommendation' in line.lower():
                in_recommendations_section = True
                continue
            elif in_recommendations_section and line.strip().startswith(('•', '-', '*', '1.', '2.', '3.')):
                recommendations.append(line.strip())
            elif in_recommendations_section and line.strip() == '':
                continue
            elif in_recommendations_section and not line.strip().startswith(('•', '-', '*')) and len(recommendations) > 0:
                break
        
        return recommendations[:3]  # Limit to 3 recommendations
    
    def _calculate_summary_confidence(self, analysis_results: Dict[str, Any]) -> float:
        """Calculate confidence score for summary."""
        if not analysis_results:
            return 0.0
        
        confidence_scores = []
        for agent_type, results in analysis_results.items():
            if isinstance(results, dict) and 'confidence_score' in results:
                confidence_scores.append(results['confidence_score'])
        
        return np.mean(confidence_scores) if confidence_scores else 0.5
    
    def _summarize_input_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Create summary of input data for metadata."""
        summary = {
            'total_keys': len(data),
            'has_analysis_results': any('analysis_type' in str(v) for v in data.values()),
            'data_types': list(set(type(v).__name__ for v in data.values()))
        }
        
        return summary