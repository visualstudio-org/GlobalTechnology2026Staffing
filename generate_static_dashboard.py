"""
Generate Static HTML Dashboard for GitHub Pages
Converts the Streamlit dashboard to a static HTML file with the same visualizations
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import json

# First Advantage Brand Colors
FA_GREEN = "#00a84f"
FA_GREEN_DARK = "#006838"
FA_GREEN_LIGHT = "#4dc47d"
FA_NAVY = "#1a1a2e"
FA_GRAY = "#6b7280"
FA_LIGHT_GRAY = "#f8f9fa"
FA_WARNING = "#f5a623"

# File path
FILE_PATH = r'C:\Users\Eric.Jaffe\OneDrive - First Advantage Corporation\2026 Budget\Global Technology 2026 Staffing Rampup Plan 011226 v2.0.xlsx'

def load_data():
    """Load data from Excel file"""
    print("Loading data from Excel...")
    
    # Load summary data
    summary_df = pd.read_excel(FILE_PATH, sheet_name='Technology Staffing Summary', header=1)
    if '#' in summary_df.columns:
        summary_df = summary_df[summary_df['#'].notna()]
    
    # Limit to first 6 rows to avoid double-counting (Excel has summary rows at bottom)
    summary_df = summary_df.head(6)
    
    # Convert numeric columns
    numeric_cols = ['# of New Roles', 'Est. Investment', 'Open Roles', 'Closed Roles']
    for col in numeric_cols:
        if col in summary_df.columns:
            summary_df[col] = pd.to_numeric(summary_df[col], errors='coerce').fillna(0)
    
    # Load detailed data
    detailed_df = pd.read_excel(FILE_PATH, sheet_name='Detailed 2026 Staffing Plans', header=1, skiprows=[0])
    
    print(f"Loaded {len(summary_df)} summary rows and {len(detailed_df)} detailed rows")
    return summary_df, detailed_df

def create_gauge_chart(value, total, title):
    """Create a gauge chart"""
    percentage = (value / total * 100) if total > 0 else 0
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={'text': title, 'font': {'size': 18, 'color': FA_NAVY}},
        number={'suffix': f' of {total}', 'font': {'size': 22, 'color': FA_GREEN}},
        gauge={
            'axis': {'range': [None, total], 'tickwidth': 1, 'tickcolor': FA_GRAY},
            'bar': {'color': FA_GREEN},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': FA_LIGHT_GRAY,
            'steps': [
                {'range': [0, total], 'color': FA_LIGHT_GRAY}
            ],
            'threshold': {
                'line': {'color': FA_WARNING, 'width': 4},
                'thickness': 0.75,
                'value': total
            }
        }
    ))
    
    fig.update_layout(
        height=300,
        margin=dict(l=60, r=60, t=80, b=40),
        paper_bgcolor='white',
        font={'family': 'Inter, sans-serif'}
    )
    
    return fig.to_html(include_plotlyjs=False, div_id=title.replace(' ', '_'), config={'displayModeBar': False})

def create_pie_chart(data, title):
    """Create a pie chart"""
    fig = px.pie(
        values=data.values,
        names=data.index,
        title=title,
        color_discrete_sequence=[FA_GREEN, FA_WARNING]
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        marker=dict(line=dict(color='white', width=2))
    )
    
    fig.update_layout(
        height=400,
        margin=dict(l=20, r=20, t=50, b=20),
        paper_bgcolor='white',
        font={'family': 'Inter, sans-serif', 'size': 14},
        title_font={'size': 18, 'color': FA_NAVY}
    )
    
    return fig.to_html(include_plotlyjs=False, div_id=title.replace(' ', '_'))

def create_bar_chart(df, x, y, title, color=FA_GREEN):
    """Create a bar chart"""
    fig = px.bar(
        df,
        x=x,
        y=y,
        title=title,
        color_discrete_sequence=[color]
    )
    
    fig.update_traces(
        marker_line_color='white',
        marker_line_width=1.5
    )
    
    fig.update_layout(
        height=400,
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor='white',
        plot_bgcolor='white',
        font={'family': 'Inter, sans-serif', 'size': 12},
        title_font={'size': 18, 'color': FA_NAVY},
        xaxis={'tickangle': -45, 'gridcolor': FA_LIGHT_GRAY},
        yaxis={'gridcolor': FA_LIGHT_GRAY}
    )
    
    return fig.to_html(include_plotlyjs=False, div_id=title.replace(' ', '_'))

def create_scatter_chart(df, x, y, size, title):
    """Create a scatter plot"""
    fig = px.scatter(
        df,
        x=x,
        y=y,
        size=size,
        title=title,
        color_discrete_sequence=[FA_GREEN],
        size_max=30
    )
    
    fig.update_traces(
        marker=dict(
            line=dict(width=1, color='white'),
            opacity=0.7
        )
    )
    
    fig.update_layout(
        height=400,
        margin=dict(l=20, r=20, t=50, b=80),
        paper_bgcolor='white',
        plot_bgcolor='white',
        font={'family': 'Inter, sans-serif', 'size': 12},
        title_font={'size': 18, 'color': FA_NAVY},
        xaxis={'tickangle': -45, 'gridcolor': FA_LIGHT_GRAY},
        yaxis={'gridcolor': FA_LIGHT_GRAY}
    )
    
    return fig.to_html(include_plotlyjs=False, div_id=title.replace(' ', '_'))

def generate_html(summary_df, detailed_df):
    """Generate the complete HTML dashboard"""
    
    print("Calculating metrics...")
    
    # Calculate overall metrics from detailed data for accurate counts
    if 'Status' in detailed_df.columns:
        open_roles = len(detailed_df[detailed_df['Status'] != 'Closed'])
        closed_roles = len(detailed_df[detailed_df['Status'] == 'Closed'])
        total_roles = len(detailed_df)
    else:
        # Fallback to summary data
        total_roles = int(summary_df['# of New Roles'].sum())
        open_roles = int(summary_df['Open Roles'].sum())
        closed_roles = int(summary_df['Closed Roles'].sum())
    
    total_investment = summary_df['Est. Investment'].sum()
    avg_cost = total_investment / total_roles if total_roles > 0 else 0
    fill_rate = (closed_roles / total_roles * 100) if total_roles > 0 else 0
    
    print("Creating charts...")
    
    # Create recruitment status pie chart
    status_data = pd.Series({
        'Open (Recruiting)': open_roles,
        'Closed (Filled)': closed_roles
    })
    
    recruitment_pie = create_pie_chart(status_data, 'Recruitment Status')
    
    # Create roles filled gauge
    gauge_html = create_gauge_chart(closed_roles, total_roles, f'Roles Filled ({fill_rate:.1f}%)')
    
    # Technology areas bar chart
    tech_bar = create_bar_chart(
        summary_df.head(10),
        'Technology Area',
        '# of New Roles',
        'Roles by Technology Area (Top 10)',
        FA_GREEN
    )
    
    # Investment bar chart
    investment_bar = create_bar_chart(
        summary_df.head(10),
        'Technology Area',
        'Est. Investment',
        'Investment by Technology Area (Top 10)',
        FA_GREEN_DARK
    )
    
    # Investment scatter plot
    investment_scatter = create_scatter_chart(
        summary_df,
        'Technology Area',
        'Est. Investment',
        '# of New Roles',
        'Investment vs. Technology Area'
    )
    
    print("Generating HTML...")
    
    # Generate HTML
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>First Advantage | Global Technology 2026 Staffing Dashboard</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Inter', sans-serif;
            background-color: #ffffff;
            color: {FA_NAVY};
            line-height: 1.6;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        header {{
            background: linear-gradient(135deg, {FA_GREEN} 0%, {FA_GREEN_DARK} 100%);
            color: white;
            padding: 30px 0;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .header-content {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 0 20px;
            display: flex;
            align-items: center;
            gap: 20px;
        }}
        
        .logo {{
            font-size: 24px;
            font-weight: 700;
        }}
        
        h1 {{
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 10px;
        }}
        
        h2 {{
            font-size: 1.8rem;
            color: {FA_NAVY};
            margin: 30px 0 20px 0;
            padding-bottom: 10px;
            border-bottom: 3px solid {FA_GREEN};
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .metric-card {{
            background: white;
            border: 2px solid {FA_LIGHT_GRAY};
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            transition: all 0.3s ease;
        }}
        
        .metric-card:hover {{
            box-shadow: 0 4px 12px rgba(0,168,79,0.15);
            border-color: {FA_GREEN};
        }}
        
        .metric-label {{
            font-size: 0.9rem;
            font-weight: 600;
            color: {FA_GRAY};
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }}
        
        .metric-value {{
            font-size: 2.5rem;
            font-weight: 700;
            color: {FA_GREEN};
            line-height: 1;
        }}
        
        .metric-delta {{
            font-size: 0.9rem;
            font-weight: 600;
            margin-top: 8px;
            color: {FA_GREEN};
        }}
        
        .chart-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 30px;
            margin-bottom: 30px;
        }}
        
        .chart-container {{
            background: white;
            border: 2px solid {FA_LIGHT_GRAY};
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }}
        
        .full-width {{
            grid-column: 1 / -1;
        }}
        
        .last-updated {{
            text-align: center;
            padding: 20px;
            color: {FA_GRAY};
            font-size: 0.9rem;
            border-top: 2px solid {FA_LIGHT_GRAY};
            margin-top: 40px;
        }}
        
        @media (max-width: 768px) {{
            h1 {{
                font-size: 1.8rem;
            }}
            
            .metrics-grid,
            .chart-grid {{
                grid-template-columns: 1fr;
            }}
            
            .metric-value {{
                font-size: 2rem;
            }}
        }}
    </style>
</head>
<body>
    <header>
        <div class="header-content">
            <div class="logo">FirstAdvantage</div>
            <div>
                <h1>Global Technology 2026 Staffing Dashboard</h1>
            </div>
        </div>
    </header>
    
    <div class="container">
        <h2>📊 Overall Metrics</h2>
        
        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-label">Total New Roles</div>
                <div class="metric-value">{total_roles:,}</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">${total_investment:,.0f}</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Open Roles</div>
                <div class="metric-value">{open_roles:,}</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Closed Roles</div>
                <div class="metric-value">{closed_roles:,}</div>
                <div class="metric-delta">↑ {fill_rate:.1f}%</div>
            </div>
            
            <div class="metric-card">
                <div class="metric-label">Avg Cost/Role</div>
                <div class="metric-value">${avg_cost:,.0f}</div>
            </div>
        </div>
        
        <h2>📈 Hiring Progress Overview</h2>
        
        <div class="chart-grid">
            <div class="chart-container">
                {gauge_html}
            </div>
            
            <div class="chart-container">
                {recruitment_pie}
            </div>
        </div>
        
        <h2>🎯 Technology Areas</h2>
        
        <div class="chart-grid">
            <div class="chart-container">
                {tech_bar}
            </div>
            
            <div class="chart-container">
                {investment_bar}
            </div>
        </div>
        
        <h2>💰 Investment Analysis</h2>
        
        <div class="chart-grid">
            <div class="chart-container full-width">
                {investment_scatter}
            </div>
        </div>
        
        <div class="last-updated">
            Last Updated: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}
        </div>
    </div>
</body>
</html>
"""
    
    return html_content

def main():
    """Main function"""
    print("=" * 60)
    print("GENERATING STATIC HTML DASHBOARD")
    print("=" * 60)
    
    try:
        # Load data
        summary_df, detailed_df = load_data()
        
        # Generate HTML
        html_content = generate_html(summary_df, detailed_df)
        
        # Write to file
        output_file = 'index.html'
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"\n✅ SUCCESS! Dashboard generated: {output_file}")
        print("\nNext steps:")
        print("1. Copy index.html to your GitHub repo")
        print("2. Commit and push to GitHub")
        print("3. Enable GitHub Pages in repo settings")
        print("4. Your dashboard will be live at: https://[username].github.io/[repo-name]/")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
