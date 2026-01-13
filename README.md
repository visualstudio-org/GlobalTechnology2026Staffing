# Global Technology 2026 Staffing Dashboard

Interactive Streamlit dashboard for tracking First Advantage Global Technology 2026 staffing ramp-up plans.

## Features

- 📊 **Interactive Visualizations** - Real-time charts and KPIs
- 🎯 **Technology Area Breakdowns** - Detailed views by tech area
- 💰 **Investment Analysis** - Track budget and costs
- 📋 **Detailed Staffing Data** - View all roles with filtering
- 📥 **Excel Export** - Download formatted reports with FA branding
- 🔄 **Auto-Refresh** - Data updates every 15 minutes
- 🎨 **First Advantage Branding** - Corporate green theme

## Setup

### Requirements

- Python 3.14+
- Required packages (see `dashboard_requirements.txt`)

### Installation

```bash
pip install -r dashboard_requirements.txt
```

### Running Locally

```bash
streamlit run staffing_dashboard.py
```

The dashboard will open at `http://localhost:8501`

### Network Access

To share on your network, the dashboard is accessible at:
```
http://<YOUR_COMPUTER_NAME>:8501
```

## Data Source

The dashboard reads from an Excel file with two sheets:
- **Technology Staffing Summary** - High-level metrics by technology area
- **Detailed 2026 Staffing Plans** - Individual role details

## Excel Export Features

Downloaded Excel files include:
- ✅ First Advantage green headers
- ✅ Auto-filter dropdowns
- ✅ Thin borders around all cells
- ✅ Auto-adjusted column widths
- ✅ Date formatting (MM/DD/YYYY)
- ✅ Timestamp-free column headers

## Dashboard Tabs

1. **Overview** - Key metrics and hiring progress
2. **Technology Areas** - Breakdown by tech area
3. **Investment Analysis** - Budget and cost analysis
4. **Detailed Data** - Full staffing data with filters

## Notes

- Dashboard is read-only - edit the Excel file directly for updates
- Excel file must be closed for dashboard to read data
- Auto-refresh runs every 15 minutes

## License

Internal First Advantage tool - Not for public distribution
