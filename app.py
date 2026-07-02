import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import urllib.parse
import io
import shutil
import os

# ============================================================
# CONFIG
# ============================================================
DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "alllinks.xlsx"
SHEET_NAME = "Links"
COLUMNS = [
    "Category", "Application Name", "Description", "Link Type", "URL", "Owner",
    "Uploaded By", "Uploaded Date", "Last Modified By", "Last Modified Date",
    "Status", "Update Available", "Version", "Department", "Priority", "Remarks"
]
STATUS_COLORS = {"Active": "🟢", "In Progress": "🔵", "Pending": "🟠", "Retired": "⚪", "Issue": "🔴"}
PRIORITY_COLORS = {"High": "🔴", "Medium": "🟡", "Low": "🟢"}
TYPE_ICONS = {
    "Power BI": "📊", "SharePoint": "🔷", "Excel": "📗", "Web Page": "🌐",
    "Application": "⚙️", "Document": "📄", "PPT / Deck": "📽️", "Ops Resource": "🔧", "Other": "🔗"
}

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Centralized Access Hub",
    page_icon="📊",
    layout="wide",
