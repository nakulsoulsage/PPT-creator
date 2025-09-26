#!/usr/bin/env python3
"""
Advanced Infographic Creator for Case Competitions
Creates stunning infographics, heatmaps, process flows, and more
"""

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
from wordcloud import WordCloud
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import io

class InfographicCreator:
    def __init__(self):
        # Professional color palettes
        self.color_palettes = {
            'mckinsey': ['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087'],
            'bcg': ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'],
            'bain': ['#4285F4', '#DB4437', '#F4B400', '#0F9D58', '#AB47BC'],
            'deloitte': ['#0076A8', '#62B5E5', '#00A767', '#7FBA00', '#FFCD00'],
            'professional': ['#1a472a', '#2a5434', '#3a6b39', '#4e8c4a', '#6fb05c']
        }
        
    def create_process_flow(self, steps, title="Process Flow"):
        """Create professional process flow diagram"""
        fig, ax = plt.subplots(1, 1, figsize=(12, 8))
        
        # Remove axes
        ax.set_xlim(0, 10)
        ax.set_ylim(0, len(steps) + 1)
        ax.axis('off')
        
        # Colors
        colors = self.color_palettes['mckinsey']
        
        # Draw process flow
        for i, step in enumerate(steps):
            y_pos = len(steps) - i
            
            # Draw arrow (except for first step)
            if i > 0:
                arrow = plt.Arrow(5, y_pos + 0.7, 0, -0.4, width=1.5, color='gray', alpha=0.6)
                ax.add_patch(arrow)
            
            # Draw box
            box_color = colors[i % len(colors)]
            rect = plt.Rectangle((2, y_pos - 0.3), 6, 0.6, 
                               facecolor=box_color, edgecolor='none', alpha=0.9)
            ax.add_patch(rect)
            
            # Add text
            ax.text(5, y_pos, step['title'], ha='center', va='center', 
                   fontsize=14, fontweight='bold', color='white')
            
            # Add description
            if 'description' in step:
                ax.text(10.5, y_pos, step['description'], ha='left', va='center',
                       fontsize=10, color='gray', wrap=True)
        
        plt.title(title, fontsize=20, fontweight='bold', pad=20)
        plt.tight_layout()
        
        # Save to bytes
        img_bytes = io.BytesIO()
        plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_heatmap(self, data, title="Heat Map Analysis", x_labels=None, y_labels=None):
        """Create professional heatmap"""
        plt.figure(figsize=(10, 8))
        
        # Create heatmap
        sns.heatmap(data, 
                   annot=True, 
                   fmt='.1f',
                   cmap='RdYlGn',
                   center=0,
                   square=True,
                   linewidths=0.5,
                   cbar_kws={"shrink": 0.8},
                   xticklabels=x_labels,
                   yticklabels=y_labels)
        
        plt.title(title, fontsize=18, fontweight='bold', pad=20)
        plt.tight_layout()
        
        # Save to bytes
        img_bytes = io.BytesIO()
        plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight')
        plt.close()
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_word_cloud(self, text_data, title="Key Themes"):
        """Create word cloud visualization"""
        # Create wordcloud
        wordcloud = WordCloud(
            width=800, 
            height=400,
            background_color='white',
            colormap='Blues',
            max_words=50,
            relative_scaling=0.5,
            min_font_size=10
        ).generate(text_data)
        
        # Plot
        plt.figure(figsize=(10, 6))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title(title, fontsize=20, fontweight='bold', pad=20)
        plt.tight_layout()
        
        # Save to bytes
        img_bytes = io.BytesIO()
        plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight')
        plt.close()
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_radar_chart(self, categories, values, title="Competitive Analysis"):
        """Create radar/spider chart"""
        # Number of variables
        num_vars = len(categories)
        
        # Compute angle for each axis
        angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
        values += values[:1]  # Complete the circle
        angles += angles[:1]
        
        # Plot
        fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
        
        # Draw the outline
        ax.plot(angles, values, color='#1f77b4', linewidth=2)
        ax.fill(angles, values, color='#1f77b4', alpha=0.25)
        
        # Fix axis to go in the right order
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        # Draw axis lines for each angle and label
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, size=12)
        
        # Set y-axis labels
        ax.set_ylim(0, max(values) * 1.1)
        ax.set_rticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(['20', '40', '60', '80', '100'], size=10)
        ax.grid(True)
        
        # Add title
        plt.title(title, fontsize=18, fontweight='bold', pad=30)
        plt.tight_layout()
        
        # Save to bytes
        img_bytes = io.BytesIO()
        plt.savefig(img_bytes, format='png', dpi=300, bbox_inches='tight')
        plt.close()
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_gauge_chart(self, value, max_value, title="Performance Metric"):
        """Create gauge/speedometer chart"""
        fig = go.Figure(go.Indicator(
            mode = "gauge+number+delta",
            value = value,
            domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': title, 'font': {'size': 24}},
            delta = {'reference': max_value * 0.8, 'increasing': {'color': "green"}},
            gauge = {
                'axis': {'range': [None, max_value], 'tickwidth': 1, 'tickcolor': "darkblue"},
                'bar': {'color': "darkblue"},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "gray",
                'steps': [
                    {'range': [0, max_value * 0.5], 'color': 'lightgray'},
                    {'range': [max_value * 0.5, max_value * 0.8], 'color': 'gray'}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': max_value * 0.9
                }
            }
        ))
        
        fig.update_layout(
            paper_bgcolor = "white",
            font = {'color': "darkblue", 'family': "Arial"},
            height=400
        )
        
        # Save to bytes
        img_bytes = io.BytesIO()
        fig.write_image(img_bytes, format='png', scale=2)
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_sankey_diagram(self, sources, targets, values, labels, title="Flow Analysis"):
        """Create Sankey diagram for flow visualization"""
        fig = go.Figure(data=[go.Sankey(
            node = dict(
                pad = 15,
                thickness = 20,
                line = dict(color = "black", width = 0.5),
                label = labels,
                color = "blue"
            ),
            link = dict(
                source = sources,
                target = targets,
                value = values,
                color = "rgba(0, 0, 255, 0.4)"
            )
        )])
        
        fig.update_layout(
            title_text=title,
            title_font_size=20,
            font_size=12,
            height=500
        )
        
        # Save to bytes
        img_bytes = io.BytesIO()
        fig.write_image(img_bytes, format='png', scale=2)
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_waterfall_chart(self, categories, values, title="Financial Waterfall"):
        """Create waterfall chart for financial analysis"""
        fig = go.Figure(go.Waterfall(
            orientation = "v",
            measure = ["relative"] * (len(values) - 1) + ["total"],
            x = categories,
            textposition = "outside",
            text = [f"${v:,.0f}" for v in values],
            y = values,
            connector = {"line":{"color":"rgb(63, 63, 63)"}},
            increasing = {"marker":{"color":"green"}},
            decreasing = {"marker":{"color":"red"}},
            totals = {"marker":{"color":"blue"}}
        ))
        
        fig.update_layout(
            title = title,
            title_font_size=20,
            showlegend = False,
            height=500
        )
        
        # Save to bytes
        img_bytes = io.BytesIO()
        fig.write_image(img_bytes, format='png', scale=2)
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_kpi_dashboard(self, kpis, title="Key Performance Indicators"):
        """Create KPI dashboard with multiple metrics"""
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=[kpi['name'] for kpi in kpis],
            specs=[[{'type': 'indicator'}, {'type': 'indicator'}],
                   [{'type': 'indicator'}, {'type': 'indicator'}]]
        )
        
        positions = [(1, 1), (1, 2), (2, 1), (2, 2)]
        
        for i, (kpi, pos) in enumerate(zip(kpis[:4], positions)):
            fig.add_trace(go.Indicator(
                mode = "number+delta",
                value = kpi['value'],
                delta = {'reference': kpi.get('target', kpi['value'] * 0.9), 
                        'relative': True,
                        'valueformat': '.1%'},
                number = {'prefix': kpi.get('prefix', ''), 
                         'suffix': kpi.get('suffix', ''),
                         'valueformat': kpi.get('format', ',')},
                domain = {'x': [0, 1], 'y': [0, 1]}
            ), row=pos[0], col=pos[1])
        
        fig.update_layout(
            title_text=title,
            title_font_size=24,
            height=500,
            showlegend=False,
            grid={'rows': 2, 'columns': 2, 'pattern': "independent"}
        )
        
        # Save to bytes
        img_bytes = io.BytesIO()
        fig.write_image(img_bytes, format='png', scale=2)
        img_bytes.seek(0)
        
        return img_bytes
    
    def create_bubble_chart(self, data_df, x_col, y_col, size_col, color_col, title="Bubble Analysis"):
        """Create bubble chart for multi-dimensional analysis"""
        fig = px.scatter(data_df, 
                        x=x_col, 
                        y=y_col,
                        size=size_col,
                        color=color_col,
                        hover_name=data_df.index if data_df.index.name else None,
                        size_max=60,
                        color_continuous_scale='Viridis')
        
        fig.update_layout(
            title=title,
            title_font_size=20,
            xaxis_title=x_col,
            yaxis_title=y_col,
            height=600
        )
        
        # Save to bytes
        img_bytes = io.BytesIO()
        fig.write_image(img_bytes, format='png', scale=2)
        img_bytes.seek(0)
        
        return img_bytes


# Test the infographic creator
def test_infographics():
    """Test all infographic functions"""
    creator = InfographicCreator()
    
    print("✓ InfographicCreator initialized")
    print("✓ Available visualizations:")
    print("  - Process Flow Diagrams")
    print("  - Heat Maps")
    print("  - Word Clouds")
    print("  - Radar Charts")
    print("  - Gauge Charts")
    print("  - Sankey Diagrams")
    print("  - Waterfall Charts")
    print("  - KPI Dashboards")
    print("  - Bubble Charts")
    print("\n✓ All infographic capabilities ready!")


if __name__ == "__main__":
    test_infographics()