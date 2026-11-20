import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import json
import os
import io
import base64
import hashlib
import re 
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from google import genai
import openpyxl  # <-- MUTLAK DIPERLUKAN UNTUK BYPASS ERROR EXCEL

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ============================================================
# CONFIG & STYLE GLOBAL
# ============================================================
LOGO_FILENAME = "logo.png"

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

LOGO_B64 = get_base64_image(LOGO_FILENAME)

st.set_page_config(page_title="FINORA AI - Pro Audit", page_icon=LOGO_FILENAME if os.path.exists(LOGO_FILENAME) else "F", layout="wide")

st.markdown("""
<style>
:root { --finora-pink: #FF5C8B; --finora-cream: #FFF5E8; --finora-rose: #F8C8D4; --finora-burgundy: #7A1F3D; --finora-text: #2D1B24; }
.stApp { background-color: var(--finora-cream); color: var(--finora-text); }
[data-testid="stSidebar"] { background-color: #FFF0F4; }
.stButton > button { background-color: var(--finora-pink); color: white !important; border-radius: 12px; border: none; font-weight: 600; padding: 10px 18px; }
.stButton > button:hover { background-color: var(--finora-burgundy); color: white !important; }
h1, h2, h3 { color: var(--finora-burgundy); }
.finora-header { background: linear-gradient(135deg, #FF5C8B 0%, #7A1F3D 100%); color: white; padding: 24px; border-radius: 16px; margin-bottom: 20px; }
.finora-header h1 { color: white; margin: 0; font-size: 30px; display: flex; align-items: center; }
.finora-header p  { color: #FFF0F4; margin: 6px 0 0 0; }
.metric-card { background: white; padding: 16px; border-radius: 12px; border-left: 4px solid #FF5C8B; box-shadow: 0 2px 8px rgba(0,0,0,0.05); margin-bottom: 15px; }
.metric-card p { color: #555; font-size: 14px; margin: 0; font-weight: 600; }
.metric-card h3 { color: #7A1F3D; margin: 5px 0 0 0; font-size: 24px; font-weight: 700; }
.insight-box { background-color: white; padding: 20px; border-radius: 10px; border-left: 5px solid #FF5C8B; margin-bottom: 15px; box-shadow: 0 2px 10px rgba(0,0,0,0.03); }
.insight-box h4 { margin-top: 0; color: #7A1F3D; font-size: 18px; margin-bottom: 10px; }
.insight-box p { margin-bottom: 8px; font-size: 14px; line-height: 1.6; color: #333; }
</style>
""", unsafe_allow_html=True)

# Mapping konsisten: nama modul audit → key flagged_data
MODULE_TO_KEY = {
    "Kelengkapan Data": "Kelengkapan Kosong",
    "Data Duplikat": "Duplikasi Ekstrem",
    "Observasi Nilai Negatif": "Nilai Negatif",
    "Overbudget (Melebihi Pagu)": "Overbudget",
    "Transaksi Weekend": "Transaksi Weekend",
    "Kata Kunci Mencurigakan": "Kata Kunci",
    "Anomali Nominal (Outlier)": "Outlier IQR",
}

def format_rupiah(val):
    try:
        val = float(val)
    except Exception:
        val = 0.0
    neg = val < 0
    v = abs(val)
    if v >= 1_000_000_000_000: s = f"Rp {v/1_000_000_000_000:,.2f} T"
    elif v >= 1_000_000_000:   s = f"Rp {v/1_000_000_000:,.2f} M"
    elif v >= 1_000_000:       s = f"Rp {v/1_000_000:,.2f} Jt"
    else:                      s = f"Rp {v:,.0f}"
    return f"({s})" if neg else s

def clean_currency(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float)): return float(x)
    x = str(x).strip()
    is_negative = False
    if x.startswith('(') and x.endswith(')'):
        is_negative = True
        x = x[1:-1]
    x = re.sub(r'[^\d,.-]', '', x)
    if not x: return 0.0
    if ',' in x and '.' in x:
        if x.rfind('.') > x.rfind(','): x = x.replace(',', '')
        else: x = x.replace('.', '').replace(',', '.')
    elif ',' in x:
        if len(x.split(',')[-1]) == 3: x = x.replace(',', '')
        else: x = x.replace(',', '.')
