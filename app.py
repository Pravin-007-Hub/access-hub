import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, date
import urllib.parse
import io
import shutil

# ============================================================
# CONFIG
# ============================================================
DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "alllinks.xlsx"
TRACKER_FILE = DATA_DIR / "daily_tracker.xlsx"
SHEET_NAME = "Links"
TRACKER_SHEET = "Tracker"
COLUMNS = [
    "Category", "Application Name", "Description", "Link Type", "URL", "Owner",
    "Uploaded By", "Uploaded Date", "Last Modified By", "Last Modified Date",
    "Status", "Update Available", "Version", "Department", "Priority", "Remarks"
]
TRACKER_COLUMNS = [
    "Task Name", "Owner", "Start Date", "Expected End Date", "ETA",
    "Progress", "Status", "Priority", "Remarks", "Created By", "Created Date",
    "Last Modified By", "Last Modified Date"
]
STATUS_COLORS = {"Active": "🟢", "In Progress": "🔵", "Pending": "🟠", "Retired": "⚪", "Issue": "🔴"}
PRIORITY_COLORS = {"High": "🔴", "Medium": "🟡", "Low": "🟢"}
TYPE_ICONS = {
    "Power BI": "📊", "SharePoint": "🔷", "Excel": "📗", "Web Page": "🌐",
    "Application": "⚙️", "Document": "📄", "PPT / Deck": "📽️", "Ops Resource": "🔧", "Other": "🔗"
}
PROGRESS_COLORS = {
    "0%": "#EF4444", "10%": "#F97316", "20%": "#F97316", "30%": "#F59E0B",
    "40%": "#F59E0B", "50%": "#EAB308", "60%": "#84CC16", "70%": "#84CC16",
    "80%": "#22C55E", "90%": "#22C55E", "100%": "#059669"
}

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="Centralized Application Access Hub", page_icon="📊", layout="wide")

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
    #MainMenu, footer, header {visibility: hidden;}
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}

    /* Header */
    .top-header {
        background: linear-gradient(135deg, #1E3A5F, #2563EB, #3B82F6);
        padding: 16px 28px; border-radius: 14px; margin-bottom: 14px;
        display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap;
        box-shadow: 0 4px 20px rgba(37,99,235,0.2);
    }
    .top-header .title-area { text-align: center; flex: 1; }
    .top-header h1 { font-size: 22px; font-weight: 800; color: #fff; margin: 0; letter-spacing: -0.01em; }
    .top-header .subtitle { color: #BFDBFE; font-size: 12px; margin: 2px 0 0; }
    .top-header .user-info {
        text-align: right; color: #BFDBFE; font-size: 11px; min-width: 180px;
    }
    .top-header .user-info b { color: #fff; font-size: 13px; }
    .top-header .logo {
        width: 42px; height: 42px; border-radius: 10px; min-width: 42px;
        background: rgba(255,255,255,0.15); display: flex; align-items: center;
        justify-content: center; font-size: 22px; backdrop-filter: blur(4px);
    }

    /* Stat cards — each a different pastel */
    .stat-row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; }
    .stat-card {
        flex: 1; min-width: 130px; padding: 14px 16px; border-radius: 12px;
        text-align: center; border: 1px solid rgba(0,0,0,0.06);
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .stat-card .num { font-size: 28px; font-weight: 800; line-height: 1.1; }
    .stat-card .lbl { font-size: 10px; color: #475569; text-transform: uppercase; letter-spacing: 0.06em; margin-top: 2px; }
    .sc-blue { background: #EFF6FF; } .sc-blue .num { color: #2563EB; }
    .sc-green { background: #ECFDF5; } .sc-green .num { color: #059669; }
    .sc-red { background: #FEF2F2; } .sc-red .num { color: #DC2626; }
    .sc-purple { background: #F5F3FF; } .sc-purple .num { color: #7C3AED; }
    .sc-yellow { background: #FEFCE8; } .sc-yellow .num { color: #CA8A04; }
    .sc-cyan { background: #ECFEFF; } .sc-cyan .num { color: #0891B2; }

    /* Badges */
    .badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 600; }
    .badge-active { background: #D1FAE5; color: #059669; }
    .badge-inprogress { background: #DBEAFE; color: #3B82F6; }
    .badge-pending { background: #FEF3C7; color: #D97706; }
    .badge-retired { background: #F1F5F9; color: #64748B; }
    .badge-issue { background: #FEE2E2; color: #DC2626; }
    .badge-yes { background: #FEE2E2; color: #DC2626; }
    .badge-no { background: #D1FAE5; color: #059669; }
    .badge-done { background: #D1FAE5; color: #059669; }
    .badge-todo { background: #DBEAFE; color: #3B82F6; }
    .badge-blocked { background: #FEE2E2; color: #DC2626; }

    /* Category */
    .cat-tag { display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; }
    .cat-dashboard { background: #DBEAFE; color: #1E40AF; }
    .cat-application { background: #FEF3C7; color: #92400E; }
    .cat-document { background: #FCE7F3; color: #9D174D; }
    .cat-report { background: #E0E7FF; color: #3730A3; }
    .cat-webpage { background: #CCFBF1; color: #065F46; }
    .cat-other { background: #F1F5F9; color: #64748B; }

    /* Tables */
    .htable { width: 100%; border-collapse: collapse; font-size: 12.5px; }
    .htable th {
        background: linear-gradient(135deg, #1E3A5F, #2563EB);
        color: white; padding: 10px 10px; text-align: left;
        font-size: 10px; font-weight: 600; text-transform: uppercase;
        letter-spacing: 0.04em; white-space: nowrap;
    }
    .htable td { padding: 8px 10px; border-bottom: 1px solid #E2E8F0; vertical-align: middle; }
    .htable tr:nth-child(even) { background:
