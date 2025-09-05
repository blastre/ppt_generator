import matplotlib.pyplot as plt
import pandas as pd
import os
import numpy as np
from model import ai

# Enhanced imports for Plotly
try:
    import plotly.graph_objects as go
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except ImportError:
    print("Warning: Plotly not installed. Using matplotlib fallback.")
    PLOTLY_AVAILABLE = False

def determine_chart_type(df: pd.DataFrame, question: str) -> str:
    """Simplified AI chart type determination - only bar and pie"""
    numeric_cols = len(df.select_dtypes(include=['number']).columns)
    categorical_cols = len(df.select_dtypes(include=['object', 'category']).columns)
    
    prompt = f"""Analyze this data and choose the BEST chart type:

Data Context:
- {df.shape[0]} rows, {df.shape[1]} columns
- {numeric_cols} numeric columns, {categorical_cols} categorical columns  
- Sample columns: {list(df.columns)[:5]}
- Question: {question}

Chart Options (ONLY these two):
- bar: comparing categories/groups, showing values across different items
- pie: parts of whole, proportions (use only if <10 categories and shows proportions)

Choose either 'bar' or 'pie' based on what best answers the question.
Return only: bar OR pie"""
    
    response = ai(prompt).strip().lower()
    # Only allow bar or pie
    return response if response in ['bar', 'pie'] else 'bar'

def create_chart_directory():
    """Create charts directory"""
    os.makedirs("charts", exist_ok=True)
    return "charts"

def create_plotly_chart(df, question, chart_path, chart_type=None):
    """Create charts using Plotly - bar and pie only"""
    if not PLOTLY_AVAILABLE:
        return create_enhanced_matplotlib_chart(df, question, chart_path)
        
    """if chart_type is None:
        chart_type = determine_chart_type(df, question)"""
# HARDCODED: Force specific chart types based on question/context
    if "slide 3" in question.lower() or "chart_3" in chart_path:
        chart_type = 'pie'
    elif "slide 4" in question.lower() or "chart_4" in chart_path:
        chart_type = 'bar'
    
    # Prepare data
    if len(df.columns) >= 2:
        x_col, y_col = df.columns[0], df.columns[1]
    else:
        x_col, y_col = 'Index', df.columns[0]
        df = df.copy()
        df['Index'] = range(len(df))
    
    # Limit data points for readability
    if len(df) > 25:
        df_sample = df.head(25)
    else:
        df_sample = df.copy()
    
    # Enhanced color palette
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#592E83', '#1B998B', '#ED217C', '#F7931E']
    
    try:
        fig = None
        
        if chart_type == 'bar':
            fig = px.bar(df_sample, x=x_col, y=y_col, 
                       title=question[:50] + "..." if len(question) > 50 else question,
                       color_discrete_sequence=colors)
            fig.update_layout(xaxis_tickangle=-45)
            
        elif chart_type == 'pie':
            if pd.api.types.is_numeric_dtype(df_sample[y_col]):
                top_data = df_sample.nlargest(8, y_col)
            else:
                top_data = df_sample.head(8)
            
            fig = px.pie(top_data, values=y_col, names=x_col,
                       title=question[:50] + "..." if len(question) > 50 else question,
                       color_discrete_sequence=colors)
        
        # Apply enhanced styling
        if fig:
            fig.update_layout(
                font=dict(family="Segoe UI, Arial", size=12),
                title_font=dict(size=18, family="Segoe UI Semibold"),
                plot_bgcolor='rgba(248,249,250,0.8)',
                paper_bgcolor='white',
                margin=dict(l=80, r=80, t=100, b=80),
                showlegend=True if chart_type in ['pie'] else False,
                width=1200,
                height=700,
                title_x=0.5
            )
            
            # Enhanced grid and axis styling for bar charts
            if chart_type == 'bar':
                fig.update_xaxes(showgrid=True, gridcolor='rgba(128,128,128,0.2)')
                fig.update_yaxes(showgrid=True, gridcolor='rgba(128,128,128,0.2)')
            
            # Save as high-quality PNG
            fig.write_image(chart_path, engine="kaleido", scale=2)
            return chart_path
            
    except Exception as e:
        print(f"Plotly chart creation failed: {e}, falling back to matplotlib")
        
    # Fallback to enhanced matplotlib
    return create_enhanced_matplotlib_chart(df_sample, question, chart_path, chart_type)

def create_enhanced_matplotlib_chart(df, question, chart_path, chart_type=None):
    """Enhanced matplotlib fallback - bar and pie only"""
    if chart_type is None:
        chart_type = determine_chart_type(df, question)
    
    # Set modern style
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    
    fig, ax = plt.subplots(figsize=(12, 8))
    
    if len(df.columns) >= 2:
        x_col, y_col = df.columns[0], df.columns[1]
        
        if chart_type == 'bar':
            # Enhanced bar chart
            bars = ax.bar(range(len(df)), df[y_col], 
                         color='#2E86AB', alpha=0.8, edgecolor='white', linewidth=0.7)
            
            ax.set_xticks(range(len(df)))
            ax.set_xticklabels([str(x)[:15] + '...' if len(str(x)) > 15 else str(x) 
                               for x in df[x_col]], rotation=45, ha='right')
            ax.set_ylabel(y_col, fontsize=14, fontweight='600')
            
            # Add value labels on bars
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                       f'{height:.1f}' if isinstance(height, float) else str(height),
                       ha='center', va='bottom', fontsize=10, fontweight='500')
                       
        elif chart_type == 'pie':
            # Enhanced pie chart
            if pd.api.types.is_numeric_dtype(df[y_col]):
                top_data = df.nlargest(8, y_col)
                sizes = top_data[y_col]
                labels = top_data[x_col]
            else:
                top_data = df.head(8)
                sizes = [1] * len(top_data)  # Equal sizes if non-numeric
                labels = top_data[x_col]
            
            colors_pie = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#592E83', '#1B998B', '#ED217C', '#F7931E']
            
            wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct='%1.1f%%',
                                            colors=colors_pie[:len(sizes)], startangle=90,
                                            textprops={'fontsize': 10})
            ax.set_aspect('equal')
        
        ax.set_title(question, fontsize=16, fontweight='700', pad=25)
    
    # Enhanced styling
    if chart_type == 'bar':
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
        ax.set_facecolor('#fafafa')
    
    plt.tight_layout()
    plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white', 
                edgecolor='none', transparent=False)
    plt.close()
    
    return chart_path

def generate_chart(df: pd.DataFrame, question: str, output_path: str) -> str:
    """Generate chart using Plotly with matplotlib fallback - bar and pie only"""
    chart_type = determine_chart_type(df, question)
    return create_plotly_chart(df, question, output_path, chart_type)