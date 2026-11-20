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
    elif '.' in x:
        if len(x.split('.')[-1]) == 3 and x.count('.') > 0: x = x.replace('.', '')
    try:
        val = float(x)
        return -val if is_negative else val
    except: return 0.0

def metric_card(label, value, help_text=""):
    st.markdown(f"""<div class="metric-card" title="{help_text}"><p>{label}</p><h3>{value}</h3></div>""", unsafe_allow_html=True)

def _clean_headers(raw_headers):
    cleaned = []
    for j, h in enumerate(raw_headers):
        h_str = str(h).strip()
        if h_str.lower() in ('nan', 'none', ''):
            h_str = f"Unnamed_{j}"
        orig, counter = h_str, 2
        while h_str in cleaned:
            h_str = f"{orig}_{counter}"
            counter += 1
        cleaned.append(h_str)
    return cleaned

# ============================================================
# FUNGSI MUTLAK PEMBACA EXCEL (TIDAK PAKAI PANDAS SAMA SEKALI)
# ============================================================
def read_raw_data(uploaded_file):
    """Membaca file mentah tanpa memicu pandas duplicate column error."""
    uploaded_file.seek(0)
    bytes_data = uploaded_file.read()
    
    if uploaded_file.name.endswith('xlsx'):
        file_stream = io.BytesIO(bytes_data)
        wb = openpyxl.load_workbook(file_stream, data_only=True, read_only=True)
        
        best_data = []
        max_rows = 0
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = list(ws.iter_rows(values_only=True))
            if len(data) > max_rows:
                max_rows = len(data)
                best_data = data
                
        return pd.DataFrame(best_data)
    else:
        file_stream = io.BytesIO(bytes_data)
        return pd.read_csv(file_stream, header=None, on_bad_lines='skip')

def safe_read_dataframe(uploaded, header_row=0):
    df_raw = read_raw_data(uploaded)
    if header_row >= len(df_raw):
        return df_raw
    cleaned = _clean_headers(df_raw.iloc[header_row].tolist())
    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = cleaned
    return df.reset_index(drop=True)

# ============================================================
# DATABASE
# ============================================================
DB_PATH = "finora.db"
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS files (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, file_hash TEXT UNIQUE, uploaded_at TEXT);
    CREATE TABLE IF NOT EXISTS analyses (id INTEGER PRIMARY KEY AUTOINCREMENT, file_id INTEGER, mode TEXT, selected_modules TEXT, status TEXT, summary_json TEXT, created_at TEXT);
    CREATE TABLE IF NOT EXISTS findings (id INTEGER PRIMARY KEY AUTOINCREMENT, analysis_id INTEGER, module TEXT, risk_level TEXT, title TEXT, detail TEXT, recommendation TEXT);
    """)
    conn.commit(); conn.close()
init_db()

def save_file(filename, content_bytes):
    file_hash = hashlib.md5(content_bytes).hexdigest()
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT id FROM files WHERE file_hash=?", (file_hash,))
    row = c.fetchone()
    if row:
        conn.close(); return row["id"]
    c.execute("INSERT INTO files(filename, file_hash, uploaded_at) VALUES (?,?,?)", (filename, file_hash, datetime.now().isoformat()))
    fid = c.lastrowid; conn.commit(); conn.close()
    return fid

def save_analysis(file_id, mode, selected_modules, summary):
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO analyses(file_id, mode, selected_modules, status, summary_json, created_at) VALUES (?,?,?,?,?,?)",
              (file_id, mode, json.dumps(selected_modules), "completed", json.dumps(summary, default=str), datetime.now().isoformat()))
    aid = c.lastrowid; conn.commit(); conn.close()
    return aid

def save_findings(analysis_id, findings):
    if not findings: return
    conn = get_conn(); c = conn.cursor()
    for f in findings:
        c.execute("INSERT INTO findings(analysis_id, module, risk_level, title, detail, recommendation) VALUES (?,?,?,?,?,?)",
                  (analysis_id, f["module"], f["risk_level"], f["title"], f["detail"], f["recommendation"]))
    conn.commit(); conn.close()

def fetch_analyses():
    conn = get_conn()
    df = pd.read_sql_query("SELECT a.id, f.filename, a.mode, a.created_at FROM analyses a JOIN files f ON a.file_id = f.id ORDER BY a.created_at DESC", conn)
    conn.close(); return df

def fetch_analysis_by_id(aid):
    conn = get_conn()
    df = pd.read_sql_query("SELECT a.id, a.summary_json, f.filename FROM analyses a JOIN files f ON a.file_id = f.id WHERE a.id=?", conn, params=(aid,))
    conn.close(); return df.iloc[0] if not df.empty else None

def fetch_findings(aid):
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM findings WHERE analysis_id=?", conn, params=(aid,))
    conn.close(); return df

# ============================================================
# AI ENGINE
# ============================================================
def get_gemini_client():
    api_key = ""
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key: return None
    return genai.Client(api_key=api_key)

def ask_gemini_stream(prompt, audit_context, system_role="Asisten Auditor Profesional"):
    try:
        client = get_gemini_client()
        if client is None:
            yield "⚠️ API Key Gemini belum diatur di Secrets."
            return
        model_id = "gemini-2.0-flash"
        full_prompt = f"Kamu adalah FINORA AI, {system_role}.\nJawab pertanyaan berdasarkan konteks data yang diberikan secara ringkas, padat, dan profesional.\n\n[KONTEKS DATA]\n{audit_context}\n\n[PERTANYAAN / INSTRUKSI]\n{prompt}"
        response_stream = client.models.generate_content_stream(model=model_id, contents=full_prompt)
        for chunk in response_stream:
            if getattr(chunk, "text", None): yield chunk.text
    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan: {str(e)}"

# ============================================================
# AUDIT ENGINE
# ============================================================
def auto_detect_columns(df):
    mapping = {"date": None, "amount": None, "desc": None, "budget": None, "invoice": None}
    for col in df.columns:
        c = str(col).lower()
        if not mapping["date"] and any(k in c for k in ["tanggal", "tgl", "date", "waktu", "periode"]): mapping["date"] = col
        if not mapping["amount"] and any(k in c for k in ["actual", "realisasi", "nominal", "jumlah", "total", "nilai", "amount", "harga", "debit", "kredit"]): mapping["amount"] = col
        if not mapping["budget"] and any(k in c for k in ["anggaran", "budget", "pagu"]): mapping["budget"] = col
        if not mapping["desc"] and any(k in c for k in ["deskripsi", "keterangan", "uraian", "category", "kategori", "nama", "item", "barang"]): mapping["desc"] = col
        if not mapping["invoice"] and any(k in c for k in ["invoice", "faktur", "nota", "ref", "no."]): mapping["invoice"] = col
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(20)
        if sample.empty: continue
        if not mapping["date"] and pd.to_datetime(sample, errors='coerce', dayfirst=True).notna().sum() > len(sample) * 0.5: mapping["date"] = col
        elif not mapping["amount"] and col != mapping["date"] and sample.str.replace(r'[^\d.-]', '', regex=True).str.isnumeric().sum() > len(sample) * 0.5: mapping["amount"] = col
    if not mapping["date"] and len(df.columns) > 0: mapping["date"] = df.columns[0]
    if not mapping["amount"] and len(df.columns) > 1: mapping["amount"] = df.columns[1]
    return mapping

def audit_empty(df, cols_to_check):
    findings, flagged = [], pd.DataFrame()
    masks = [df[c].isna() for c in cols_to_check if c]
    if masks:
        combined_mask = masks[0]
        for m in masks[1:]: combined_mask |= m
        flagged = df[combined_mask]
        if not flagged.empty:
            findings.append({"module": "Kelengkapan Data", "risk_level": "Tinggi", "title": f"{len(flagged)} baris memiliki data vital kosong", "detail": "Ada data kosong pada kolom Tanggal/Nominal/Deskripsi.", "recommendation": "Lengkapi data sebelum diproses."})
    return findings, flagged

def audit_duplicates(df, cols_to_check):
    findings, flagged = [], pd.DataFrame()
    if cols_to_check:
        flagged = df[df.duplicated(subset=cols_to_check, keep=False)]
        if not flagged.empty:
            findings.append({"module": "Data Duplikat", "risk_level": "Tinggi", "title": f"{len(flagged)} baris transaksi ganda (100% sama)", "detail": "Terdapat transaksi dengan nilai/deskripsi yang sama persis.", "recommendation": "Pastikan bukan input dua kali (double-entry)."})
    return findings, flagged

def audit_negative(df, amount_col):
    findings, flagged = [], pd.DataFrame()
    if amount_col:
        s = pd.to_numeric(df[amount_col], errors="coerce")
        flagged = df[s < 0]
        if not flagged.empty:
            findings.append({"module": "Observasi Nilai Negatif", "risk_level": "Rendah", "title": f"{len(flagged)} transaksi bernilai negatif", "detail": "Ditemukan nilai minus pada kolom nominal.", "recommendation": "Lakukan uji petik untuk memastikan transaksi negatif memiliki memo retur/koreksi."})
    return findings, flagged

def audit_overbudget(df, budget_col, amount_col):
    findings, flagged = [], pd.DataFrame()
    if budget_col and amount_col:
        try:
            b = pd.to_numeric(df[budget_col], errors="coerce").fillna(0)
            a = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
            flagged = df[a > b]
            if not flagged.empty:
                findings.append({"module": "Overbudget (Melebihi Pagu)", "risk_level": "Tinggi", "title": f"{len(flagged)} akun melebihi pagu anggaran", "detail": "Nilai realisasi lebih besar daripada anggaran.", "recommendation": "Minta penjelasan terkait kelebihan anggaran."})
        except: pass
    return findings, flagged

def audit_weekend(df, date_col):
    findings, flagged = [], pd.DataFrame()
    if date_col:
        try:
            dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
            flagged = df[dates.dt.dayofweek >= 5]
            if not flagged.empty:
                findings.append({"module": "Transaksi Weekend", "risk_level": "Tinggi", "title": f"{len(flagged)} transaksi di hari Sabtu/Minggu", "detail": "Transaksi di luar hari kerja.", "recommendation": "Periksa otorisasi transaksi akhir pekan."})
        except: pass
    return findings, flagged

def audit_suspicious_words(df, desc_col):
    findings, flagged = [], pd.DataFrame()
    keywords = ["adjustment", "adj", "void", "koreksi", "manual", "write-off", "hapus"]
    if desc_col:
        try:
            pattern = '|'.join(keywords)
            flagged = df[df[desc_col].astype(str).str.lower().str.contains(pattern, na=False)]
            if not flagged.empty:
                findings.append({"module": "Kata Kunci Mencurigakan", "risk_level": "Sedang", "title": f"{len(flagged)} transaksi mengandung kata kunci adjustment/void", "detail": "Transaksi manual sering menjadi celah manipulasi.", "recommendation": "Minta penjelasan tertulis dari pembuat jurnal."})
        except: pass
    return findings, flagged

def audit_outliers_iqr(df, amount_col):
    findings, flagged = [], pd.DataFrame()
    if amount_col:
        try:
            s = pd.to_numeric(df[amount_col], errors="coerce").dropna()
            if not s.empty:
                Q1, Q3 = s.quantile(0.25), s.quantile(0.75)
                IQR = Q3 - Q1
                lower, upper = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
                flagged = df[(pd.to_numeric(df[amount_col], errors="coerce") < lower) | (pd.to_numeric(df[amount_col], errors="coerce") > upper)]
                if not flagged.empty:
                    findings.append({"module": "Anomali Nominal (Outlier)", "risk_level": "Tinggi", "title": f"{len(flagged)} nilai ekstrem (di luar batas IQR)", "detail": f"Batas wajar: {lower:,.0f} s.d {upper:,.0f}", "recommendation": "Uji petik transaksi terbesar/terkecil."})
        except: pass
    return findings, flagged

# ============================================================
# SMART FINANCIAL ENGINE & INSIGHT
# ============================================================
def auto_extract_financials(df, entity_type):
    results = {"pendapatan": 0.0, "laba": 0.0, "aset_lancar": 0.0, "liabilitas_lancar": 0.0, "total_aset": 0.0, "total_liabilitas": 0.0, "ekuitas": 0.0, "real_pendapatan": 0.0, "ang_pendapatan": 0.0, "real_belanja": 0.0, "ang_belanja": 0.0, "surplus": 0.0, "pad": 0.0}
    str_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
    num_cols = df.select_dtypes(include=['number', 'float', 'int']).columns.tolist()
    if not str_cols or not num_cols: return results
    name_col = str_cols[0] 
    real_col, ang_col = None, None
    for c in num_cols:
        c_str = str(c).lower()
        if 'realisasi' in c_str or 'real' in c_str: real_col = c
        elif 'anggaran' in c_str or 'ang' in c_str or 'pagu' in c_str: ang_col = c
    if not real_col and len(num_cols) > 0: real_col = num_cols[0]
    if not ang_col and len(num_cols) > 1: ang_col = num_cols[1]
    clean_names = df[name_col].astype(str).str.lower().str.replace(r'[^a-z\s]', '', regex=True).str.strip()

    def get_total(target_col, exact_keys, contains_keys=[]):
        if not target_col: return 0.0
        if exact_keys:
            mask_exact = clean_names.isin(exact_keys)
            if mask_exact.any(): return pd.to_numeric(df.loc[mask_exact, target_col], errors='coerce').fillna(0).sum()
        if contains_keys:
            mask_contains = clean_names.str.contains('|'.join(contains_keys), na=False)
            if mask_contains.any(): return pd.to_numeric(df.loc[mask_contains, target_col], errors='coerce').fillna(0).sum()
        return 0.0

    if "Pemerintahan" in entity_type:
        real_p = get_total(real_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan daerah', 'pendapatan'])
        if real_p == 0: real_p = get_total(real_col, [], ['pendapatan asli daerah', 'dana perimbangan', 'dana transfer', 'lainlain pendapatan'])
        ang_p = get_total(ang_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan daerah', 'pendapatan'])
        if ang_p == 0: ang_p = get_total(ang_col, [], ['pendapatan asli daerah', 'dana perimbangan', 'dana transfer', 'lainlain pendapatan'])
        real_b = get_total(real_col, ['total belanja', 'jumlah belanja', 'belanja daerah', 'belanja'])
        if real_b == 0: real_b = get_total(real_col, [], ['belanja operasi', 'belanja modal', 'belanja tak terduga'])
        ang_b = get_total(ang_col, ['total belanja', 'jumlah belanja', 'belanja daerah', 'belanja'])
        if ang_b == 0: ang_b = get_total(ang_col, [], ['belanja operasi', 'belanja modal', 'belanja tak terduga'])
        pad = get_total(real_col, ['pendapatan asli daerah', 'jumlah pendapatan asli daerah', 'total pendapatan asli daerah', 'pad'], ['pendapatan asli daerah'])
        surplus = get_total(real_col, [], ['surplus', 'defisit'])
        if surplus == 0: surplus = real_p - real_b
        results.update({"real_pendapatan": real_p, "ang_pendapatan": ang_p, "real_belanja": real_b, "ang_belanja": ang_b, "pad": pad, "surplus": surplus})
    else:
        results["pendapatan"] = get_total(real_col, ['total pendapatan', 'pendapatan', 'total penjualan', 'penjualan'], ['pendapatan', 'penjualan', 'revenue'])
        results["laba"] = get_total(real_col, ['laba bersih', 'laba tahun berjalan'], ['laba bersih', 'net income'])
        results["total_aset"] = get_total(real_col, ['total aset', 'jumlah aset'], ['total aset', 'total asset'])
        results["ekuitas"] = get_total(real_col, ['total ekuitas', 'total modal'], ['total ekuitas', 'total equity'])
        results["total_liabilitas"] = get_total(real_col, ['total liabilitas', 'total utang'], ['total liabilitas', 'total utang'])
        results["aset_lancar"] = get_total(real_col, ['total aset lancar'], ['aset lancar', 'current asset'])
        results["liabilitas_lancar"] = get_total(real_col, ['total liabilitas jangka pendek', 'total utang lancar'], ['liabilitas jangka pendek', 'utang lancar', 'current liab'])
    return results

def hitung_rasio_dari_nilai(nilai, entity_type):
    rasio = {}
    if "Pemerintahan" in entity_type:
        real_p = nilai.get("real_pendapatan", 0); ang_p = nilai.get("ang_pendapatan", 0)
        real_b = nilai.get("real_belanja", 0); ang_b = nilai.get("ang_belanja", 0); pad = nilai.get("pad", 0)
        rasio["Efektivitas Pendapatan (%)"] = round((real_p / ang_p) * 100, 2) if ang_p > 0 else 0.0
        rasio["Efisiensi Belanja (%)"] = round((real_b / ang_b) * 100, 2) if ang_b > 0 else 0.0
        rasio["Kemandirian Daerah (%)"] = round((pad / real_p) * 100, 2) if real_p > 0 else 0.0
        rasio["Pertumbuhan Pendapatan (%)"] = 0.0 
    else:
        pendapatan = nilai.get("pendapatan", 0); laba = nilai.get("laba", 0)
        aset_lancar = nilai.get("aset_lancar", 0); liab_lancar = nilai.get("liabilitas_lancar", 0)
        total_aset = nilai.get("total_aset", 0); total_liab = nilai.get("total_liabilitas", 0); ekuitas = nilai.get("ekuitas", 0)
        rasio["Current Ratio (x)"] = round(aset_lancar / liab_lancar, 2) if liab_lancar > 0 else 0.0
        rasio["Net Profit Margin (%)"] = round((laba / pendapatan) * 100, 2) if pendapatan > 0 else 0.0
        rasio["ROA (%)"] = round((laba / total_aset) * 100, 2) if total_aset > 0 else 0.0
        rasio["ROE (%)"] = round((laba / ekuitas) * 100, 2) if ekuitas > 0 else 0.0
        rasio["Debt to Equity (x)"] = round(total_liab / ekuitas, 2) if ekuitas > 0 else 0.0
    return rasio

def get_detailed_insight(ratio_name, val):
    if val == 0 or pd.isna(val):
        if "Pertumbuhan" in ratio_name: return "<b>Kondisi:</b> 0% (Tidak Ada Data).<br><b>Penyebab:</b> Butuh data tahun sebelumnya.<br><b>Dampak:</b> Tidak bisa lihat tren.<br><b>Rekomendasi:</b> Simpan laporan tahun ini, bandingkan tahun depan."
        elif "Efektivitas" in ratio_name: return "<b>Kondisi:</b> Rasio 0%.<br><b>Penyebab:</b> Anggaran Pendapatan bernilai 0.<br><b>Dampak:</b> Tidak bisa ukur kinerja.<br><b>Rekomendasi:</b> Periksa kolom Anggaran Pendapatan."
        elif "Efisiensi" in ratio_name: return "<b>Kondisi:</b> Rasio 0%.<br><b>Penyebab:</b> Anggaran Belanja bernilai 0.<br><b>Dampak:</b> Tidak bisa ukur penyerapan.<br><b>Rekomendasi:</b> Pastikan kolom Anggaran Belanja valid."
        elif "Kemandirian" in ratio_name: return "<b>Kondisi:</b> 0%.<br><b>Penyebab:</b> PAD atau Pendapatan 0.<br><b>Dampak:</b> Tidak Mandiri.<br><b>Rekomendasi:</b> Lakukan ekstensifikasi pajak."
        elif "Current Ratio" in ratio_name: return "<b>Kondisi:</b> 0x.<br><b>Penyebab:</b> Aset Lancar 0.<br><b>Dampak:</b> Risiko Likuiditas Ekstrem.<br><b>Rekomendasi:</b> Tambah modal kerja."
        elif "Debt to Equity" in ratio_name: return "<b>Kondisi:</b> 0x.<br><b>Penyebab:</b> Tidak ada utang.<br><b>Dampak:</b> Sangat Aman.<br><b>Rekomendasi:</b> Bisa mulai gunakan utang produktif."
        else: return "<b>Kondisi:</b> Rasio tidak dapat dihitung.<br><b>Rekomendasi:</b> Periksa format Excel."
    if "Efektivitas" in ratio_name:
        if val >= 100: return "<b>Kondisi:</b> Sangat Baik.<br><b>Penyebab:</b> Kinerja pemungutan optimal.<br><b>Dampak:</b> Kapasitas fiskal meningkat.<br><b>Rekomendasi:</b> Pertahankan, eksplorasi PAD baru."
        elif val >= 90: return "<b>Kondisi:</b> Baik.<br><b>Penyebab:</b> Pemungutan sesuai rencana.<br><b>Dampak:</b> Keuangan stabil.<br><b>Rekomendasi:</b> Tingkatkan pengawasan."
        else: return "<b>Kondisi:</b> Kurang.<br><b>Penyebab:</b> Target terlalu tinggi / ekonomi lesu.<br><b>Dampak:</b> Defisit, pembangunan terhambat.<br><b>Rekomendasi:</b> Intensifikasi & Ekstensifikasi pajak."
    elif "Efisiensi" in ratio_name:
        if val > 100: return "<b>Kondisi:</b> Kurang Efisien (Overbudget).<br><b>Penyebab:</b> Pengeluaran tak terduga/inflasi.<br><b>Dampak:</b> Defisit, temuan audit.<br><b>Rekomendasi:</b> Evaluasi anggaran, revisi DPA."
        elif val >= 90: return "<b>Kondisi:</b> Sangat Efisien.<br><b>Penyebab:</b> Eksekusi sesuai jadwal.<br><b>Dampak:</b> Program tercapai.<br><b>Rekomendasi:</b> Pastikan kualitas hasil (Value for Money)."
        else: return "<b>Kondisi:</b> Penyerapan Rendah.<br><b>Penyebab:</b> Keterlambatan tender/kendala teknis.<br><b>Dampak:</b> Target tertunda, idle money.<br><b>Rekomendasi:</b> Percepat tender, monev berkala."
    elif "Kemandirian" in ratio_name:
        if val >= 50: return "<b>Kondisi:</b> Mandiri.<br><b>Penyebab:</b> Basis pajak kuat.<br><b>Dampak:</b> Fleksibilitas fiskal tinggi.<br><b>Rekomendasi:</b> Kembangkan BUMD."
        elif val >= 25: return "<b>Kondisi:</b> Sedang.<br><b>Penyebab:</b> Sektor unggulan belum maksimal.<br><b>Dampak:</b> Rentan kebijakan pusat.<br><b>Rekomendasi:</b> Gali potensi pariwisata/parkir."
        else: return "<b>Kondisi:</b> Rendah.<br><b>Penyebab:</b> Pajak/retribusi kecil.<br><b>Dampak:</b> Ketergantungan transfer pusat.<br><b>Rekomendasi:</b> Reformasi perpajakan daerah."
    elif "Current Ratio" in ratio_name:
        if val >= 1.5: return "<b>Kondisi:</b> Likuiditas Sehat.<br><b>Penyebab:</b> Manajemen modal kerja baik.<br><b>Dampak:</b> Kredibilitas tinggi.<br><b>Rekomendasi:</b> Gunakan kas berlebih untuk investasi produktif."
        else: return "<b>Kondisi:</b> Risiko Likuiditas.<br><b>Penyebab:</b> Kas menipis, piutang macet.<br><b>Dampak:</b> Risiko gagal bayar.<br><b>Rekomendasi:</b> Percepat penagihan piutang, negosiasi ulang utang."
    elif "Debt to Equity" in ratio_name:
        if val > 2.0: return "<b>Kondisi:</b> High Leverage.<br><b>Penyebab:</b> Agresif pakai utang.<br><b>Dampak:</b> Beban bunga tinggi.<br><b>Rekomendasi:</b> Kurangi utang, cari pendanaan ekuitas."
        else: return "<b>Kondisi:</b> Struktur Modal Sehat.<br><b>Penyebab:</b> Utang & ekuitas seimbang.<br><b>Dampak:</b> Risiko finansial rendah.<br><b>Rekomendasi:</b> Aman untuk ekspansi."
    return "<b>Analisis:</b> Detail belum tersedia."

# ============================================================
# EXPORT FUNCTIONS (PDF & EXCEL)
# ============================================================
def generate_excel(flagged_data_dict, findings_list):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_findings = pd.DataFrame(findings_list)
        if not df_findings.empty:
            df_findings.columns = ["Modul", "Tingkat Risiko", "Judul", "Detail", "Rekomendasi"]
            df_findings.to_excel(writer, sheet_name="Ringkasan Temuan", index=False)
        for modul, df_flag in flagged_data_dict.items():
            if not df_flag.empty: df_flag.to_excel(writer, sheet_name=modul[:31], index=False)
    buffer.seek(0)
    return buffer

def generate_pdf(entitas, periode, filename, score, findings_list, flagged_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='Title', parent=styles['Heading1'], textColor=colors.HexColor('#7A1F3D'), alignment=1)
    heading_style = ParagraphStyle(name='Heading', parent=styles['Heading2'], textColor=colors.HexColor('#FF5C8B'))
    normal_style = styles['Normal']
    elements.append(Paragraph("KERTAS KERJA AUDIT (KKA)", title_style))
    elements.append(Paragraph("FINORA Financial Intelligence System", normal_style))
    elements.append(Spacer(1, 10))
    header_data = [["Nama Entitas", f": {entitas if entitas else '-'}"], ["Periode Audit", f": {periode if periode else '-'}"], ["Nama File", f": {filename}"], ["Tanggal Pemeriksaan", f": {datetime.now().strftime('%d %B %Y')}"], ["Skor Risiko", f": {score}/100"]]
    t_header = Table(header_data, colWidths=[120, 350])
    t_header.setStyle(TableStyle([('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 10), ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]))
    elements.append(t_header)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("A. Ringkasan Temuan", heading_style))
    elements.append(Spacer(1, 10))
    summary_data = [["No", "Kode", "Kategori Uji", "Risiko", "Deskripsi Singkat"]]
    for idx, f in enumerate(findings_list, 1):
        summary_data.append([str(idx), f"F-{idx:02d}", f['module'], f['risk_level'], f['title']])
    t_summary = Table(summary_data, colWidths=[30, 50, 100, 60, 230])
    t_summary.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF5C8B')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 9), ('GRID', (0, 0), (-1, -1), 1, colors.grey)]))
    elements.append(t_summary)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("B. Rincian Temuan & Bukti", heading_style))
    for idx, f in enumerate(findings_list, 1):
        elements.append(Paragraph(f"<b>F-{idx:02d}: {f['title']}</b>", normal_style))
        elements.append(Paragraph(f"<b>1. Kondisi:</b> {f['detail']}", normal_style))
        elements.append(Paragraph(f"<b>2. Rekomendasi:</b> {f['recommendation']}", normal_style))
        df_flag = pd.DataFrame()
        key = MODULE_TO_KEY.get(f['module'])
        if key and key in flagged_data and not flagged_data[key].empty:
            df_flag = flagged_data[key]
        if not df_flag.empty:
            elements.append(Spacer(1, 5))
            elements.append(Paragraph("<i>Bukti Audit:</i>", normal_style))
            preview_df = df_flag.head(5)
            raw_data = [preview_df.columns.tolist()] + preview_df.values.tolist()
            raw_data = [[str(cell) for cell in row] for row in raw_data]
            t_raw = Table(raw_data)
            t_raw.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), ('FONTSIZE', (0, 0), (-1, -1), 8)]))
            elements.append(t_raw)
        elements.append(Spacer(1, 20))
    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_fin_excel(nilai, rasio, entity_type):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        data_nilai = []
        if "Pemerintahan" in entity_type: data_nilai = [{"Komponen": "Realisasi Pendapatan", "Nilai": nilai.get("real_pendapatan", 0)}, {"Komponen": "Anggaran Pendapatan", "Nilai": nilai.get("ang_pendapatan", 0)}, {"Komponen": "Realisasi Belanja", "Nilai": nilai.get("real_belanja", 0)}, {"Komponen": "Anggaran Belanja", "Nilai": nilai.get("ang_belanja", 0)}, {"Komponen": "Surplus / Defisit", "Nilai": nilai.get("surplus", 0)}, {"Komponen": "Pendapatan Asli Daerah (PAD)", "Nilai": nilai.get("pad", 0)}]
        else: data_nilai = [{"Komponen": "Total Pendapatan", "Nilai": nilai.get("pendapatan", 0)}, {"Komponen": "Laba Bersih", "Nilai": nilai.get("laba", 0)}, {"Komponen": "Total Aset", "Nilai": nilai.get("total_aset", 0)}, {"Komponen": "Total Liabilitas", "Nilai": nilai.get("total_liabilitas", 0)}, {"Komponen": "Total Ekuitas", "Nilai": nilai.get("ekuitas", 0)}, {"Komponen": "Aset Lancar", "Nilai": nilai.get("aset_lancar", 0)}, {"Komponen": "Liabilitas Jangka Pendek", "Nilai": nilai.get("liabilitas_lancar", 0)}]
        pd.DataFrame(data_nilai).to_excel(writer, sheet_name="Data Keuangan", index=False)
        data_rasio = [{"Rasio": k, "Nilai (%)": v} for k, v in rasio.items()]
        pd.DataFrame(data_rasio).to_excel(writer, sheet_name="Rasio Keuangan", index=False)
    buffer.seek(0)
    return buffer

def generate_fin_pdf(nilai, rasio, entity_type):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='Title', parent=styles['Heading1'], textColor=colors.HexColor('#7A1F3D'), alignment=1)
    heading_style = ParagraphStyle(name='Heading', parent=styles['Heading2'], textColor=colors.HexColor('#FF5C8B'))
    normal_style = styles['Normal']
    elements.append(Paragraph("LAPORAN ANALISIS KEUANGAN", title_style))
    elements.append(Paragraph(f"FINORA Financial Intelligence System - {entity_type}", normal_style))
    elements.append(Paragraph(f"Tanggal Analisis: {datetime.now().strftime('%d %B %Y')}", normal_style))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("A. Ringkasan Data Keuangan", heading_style))
    elements.append(Spacer(1, 10))
    header_data = [["Komponen", "Nilai (Rp)"]]
    if "Pemerintahan" in entity_type: header_data.extend([["Realisasi Pendapatan", f"{nilai.get('real_pendapatan', 0):,.0f}"], ["Anggaran Pendapatan", f"{nilai.get('ang_pendapatan', 0):,.0f}"], ["Realisasi Belanja", f"{nilai.get('real_belanja', 0):,.0f}"], ["Anggaran Belanja", f"{nilai.get('ang_belanja', 0):,.0f}"], ["Surplus / Defisit", f"{nilai.get('surplus', 0):,.0f}"], ["Pendapatan Asli Daerah (PAD)", f"{nilai.get('pad', 0):,.0f}"]])
    else: header_data.extend([["Total Pendapatan", f"{nilai.get('pendapatan', 0):,.0f}"], ["Laba Bersih", f"{nilai.get('laba', 0):,.0f}"], ["Total Aset", f"{nilai.get('total_aset', 0):,.0f}"], ["Total Liabilitas", f"{nilai.get('total_liabilitas', 0):,.0f}"], ["Total Ekuitas", f"{nilai.get('ekuitas', 0):,.0f}"]])
    t_data = Table(header_data, colWidths=[250, 200])
    t_data.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF5C8B')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 10), ('GRID', (0, 0), (-1, -1), 1, colors.grey), ('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'RIGHT')]))
    elements.append(t_data)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("B. Rasio Keuangan", heading_style))
    elements.append(Spacer(1, 10))
    rasio_data = [["Rasio", "Nilai"]]
    for k, v in rasio.items(): rasio_data.append([k, f"{v}"])
    t_rasio = Table(rasio_data, colWidths=[300, 150])
    t_rasio.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 10), ('GRID', (0, 0), (-1, -1), 1, colors.grey), ('ALIGN', (0, 0), (-1, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'RIGHT')]))
    elements.append(t_rasio)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("C. Analisis & Rekomendasi Mendalam", heading_style))
    elements.append(Spacer(1, 10))
    for k, v in rasio.items():
        insight_text = get_detailed_insight(k, v).replace("<br>", "<br/>")
        elements.append(Paragraph(f"<b>{k} ({v})</b>", normal_style))
        elements.append(Paragraph(insight_text, normal_style))
        elements.append(Spacer(1, 10))
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# UI COMPONENTS & HALAMAN
# ============================================================
def finora_header(title, subtitle):
    logo_html = f'<img src="data:image/png;base64,{LOGO_B64}" style="height: 40px; vertical-align: middle; margin-right: 12px; margin-top: -5px; border-radius: 8px;">' if LOGO_B64 else ''
    st.markdown(f'<div class="finora-header"><h1>{logo_html} {title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)

def render_audit_results(all_findings, flagged_data, filename, entitas, periode, score):
    color = "#C62828" if score > 50 else "#EF6C00" if score > 20 else "#2E7D32"
    st.markdown(f"""<div style="background-color: white; padding: 20px; border-radius: 10px; border: 2px solid #FF5C8B; margin-bottom: 20px;"><h2 style="color: #7A1F3D; text-align: center; margin-bottom: 5px;">KERTAS KERJA AUDIT (KKA)</h2><p style="text-align: center; color: gray; margin-top: 0;">FINORA Financial Intelligence System</p><hr><table style="width: 100%; font-size: 14px;"><tr><td><b>Nama Entitas</b></td><td>: {entitas if entitas else '-'}</td></tr><tr><td><b>Periode Audit</b></td><td>: {periode if periode else '-'}</td></tr><tr><td><b>Nama File</b></td><td>: {filename}</td></tr><tr><td><b>Tanggal Pemeriksaan</b></td><td>: {datetime.now().strftime('%d %B %Y')}</td></tr><tr><td><b>Skor Risiko</b></td><td>: <span style='color:{color}; font-weight:bold; font-size:18px;'>{score}/100</span></td></tr></table></div>""", unsafe_allow_html=True)
    tab_kka, tab_grafik, tab_ai = st.tabs(["📝 Laporan Temuan & Bukti", "📊 Visualisasi Grafik", "🤖 Tanya FINORA AI"])
    with tab_kka:
        st.markdown("### A. Ringkasan Temuan")
        summary_data = [{"No": idx, "Kode": f"F-{idx:02d}", "Kategori Uji": f["module"], "Risiko": f["risk_level"], "Deskripsi": f["title"]} for idx, f in enumerate(all_findings, 1)]
        if summary_data: st.table(pd.DataFrame(summary_data))
        else: st.info("Tidak ada temuan.")
        st.markdown("---")
        st.markdown("### B. Rincian Temuan")
        for idx, f in enumerate(all_findings, 1):
            with st.expander(f"[{f['risk_level']}] F-{idx:02d}: {f['title']}"):
                c1, c2 = st.columns([1, 2])
                with c1:
                    st.markdown(f"**1. Kondisi**\n{f['detail']}")
                    st.markdown("**2. Kriteria**\nSOP Internal / SAK/PSAK.")
                with c2:
                    st.markdown("**3. Sebab**\nKelemahan kontrol internal / human error.")
                    st.markdown("**4. Akibat**\nBerpotensi kerugian finansial / laporan tidak wajar.")
                st.error(f"**5. Rekomendasi:** {f['recommendation']}")
                key = MODULE_TO_KEY.get(f['module'])
                if key and key in flagged_data and not flagged_data[key].empty:
                    st.dataframe(flagged_data[key].head(10), use_container_width=True)
        st.markdown("---")
        st.warning(f"Ditemukan **{len(all_findings)} indikasi masalah**. Data memerlukan verifikasi lebih lanjut.")
        st.markdown("### 📥 Download Kertas Kerja Audit")
        col1, col2 = st.columns(2)
        col1.download_button("⬇️ Download Excel", data=generate_excel(flagged_data, all_findings), file_name=f"KKA_{entitas}.xlsx", use_container_width=True)
        col2.download_button("⬇️ Download PDF", data=generate_pdf(entitas, periode, filename, score, all_findings, flagged_data), file_name=f"KKA_{entitas}.pdf", use_container_width=True)
    with tab_grafik:
        if all_findings:
            c1, c2 = st.columns(2)
            with c1:
                risk_df = pd.DataFrame(all_findings)["risk_level"].value_counts().reset_index()
                st.plotly_chart(px.pie(risk_df, names="risk_level", values="count", color="risk_level", color_discrete_map={"Tinggi": "#C62828", "Sedang": "#EF6C00", "Rendah": "#2E7D32"}, hole=0.4), use_container_width=True)
            with c2:
                mod_df = pd.DataFrame(all_findings)["module"].value_counts().reset_index()
                st.plotly_chart(px.bar(mod_df, x="count", y="module", orientation='h', color_discrete_sequence=["#FF5C8B"]), use_container_width=True)
        else: st.info("Belum ada data.")
    with tab_ai:
        st.markdown("### 🤖 Tanya FINORA AI")
        context_text = f"Entitas: {entitas} | Skor: {score}/100\nTemuan:\n" + "".join([f"- [{f['risk_level']}] {f['module']} ({f['title']})\n" for f in all_findings])
        if "ai_chat_history" not in st.session_state: st.session_state.ai_chat_history = []
        for msg in st.session_state.ai_chat_history:
            with st.chat_message(msg["role"]): st.markdown(msg["content"])
        with st.form(key="ai_chat_form", clear_on_submit=True):
            prompt = st.text_input("Tulis pertanyaan audit Anda...")
            if st.form_submit_button("Kirim ke AI", use_container_width=True) and prompt:
                st.session_state.ai_chat_history.append({"role": "user", "content": prompt})
                with st.chat_message("assistant"): response = st.write_stream(ask_gemini_stream(prompt, context_text))
                st.session_state.ai_chat_history.append({"role": "assistant", "content": response})
                st.rerun()

def page_audit():
    finora_header("Audit Profesional", "Sistem deteksi otomatis anomali transaksi")
    with st.container():
        st.markdown("### 1. Upload Data")
        c1, c2 = st.columns([3, 1])
        uploaded = c1.file_uploader("Upload File Excel/CSV", type=["xlsx", "csv"], key="audit_upload")
        header_row = c2.number_input("Header di baris ke-", min_value=0, value=0, key="audit_header")
    if uploaded:
        try:
            df = safe_read_dataframe(uploaded, header_row=header_row)
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {str(e)}")
            st.stop()
        detected = auto_detect_columns(df)
        st.markdown("### 2. Pemetaan Kolom")
        st.success("✨ FINORA AI telah mendeteksi otomatis kolom.")
        with st.expander("⚙️ Koreksi Pemetaan Kolom (Klik Jika Perlu)", expanded=False):
            c1, c2 = st.columns(2)
            date_col = c1.selectbox("Kolom Tanggal", list(df.columns), index=list(df.columns).index(detected["date"]) if detected["date"] in list(df.columns) else 0)
            desc_col = c2.selectbox("Kolom Deskripsi", [""] + list(df.columns), index=list(df.columns).index(detected["desc"])+1 if detected["desc"] in list(df.columns) else 0)
            c3, c4, c5 = st.columns(3)
            amount_col = c3.selectbox("Kolom Nominal", list(df.columns), index=list(df.columns).index(detected["amount"]) if detected["amount"] in list(df.columns) else 0)
            budget_col = c4.selectbox("Kolom Anggaran", [""] + list(df.columns), index=list(df.columns).index(detected["budget"])+1 if detected["budget"] in list(df.columns) else 0)
            invoice_col = c5.selectbox("Kolom No. Faktur", [""] + list(df.columns), index=list(df.columns).index(detected["invoice"])+1 if detected["invoice"] in list(df.columns) else 0)
        st.markdown("### 3. Informasi KKA")
        c_kka1, c_kka2 = st.columns(2)
        entitas = c_kka1.text_input("Nama Entitas / Perusahaan", placeholder="Contoh: PT Maju Jaya", key="audit_entitas")
        periode = c_kka2.text_input("Periode Audit", placeholder="Contoh: Tahun Buku 2024", key="audit_periode")
        st.markdown("### 4. Jalankan Prosedur")
        if st.button("🚀 JALANKAN PROSEDUR AUDIT", use_container_width=True):
            with st.spinner("Memproses data..."):
                
                if amount_col: df[amount_col] = df[amount_col].apply(clean_currency)
                if budget_col: df[budget_col] = df[budget_col].apply(clean_currency)
                if invoice_col: df[invoice_col] = df[invoice_col].astype(str)
                
                all_findings, flagged_data = [], {}
                f, flag = audit_empty(df, [c for c in [date_col, amount_col, desc_col] if c]); all_findings.extend(f); flagged_data["Kelengkapan Kosong"] = flag
                f, flag = audit_duplicates(df, [c for c in [date_col, amount_col, desc_col, invoice_col] if c]); all_findings.extend(f); flagged_data["Duplikasi Ekstrem"] = flag
                if date_col: f, flag = audit_weekend(df, date_col); all_findings.extend(f); flagged_data["Transaksi Weekend"] = flag
                if desc_col: f, flag = audit_suspicious_words(df, desc_col); all_findings.extend(f); flagged_data["Kata Kunci"] = flag
                if amount_col:
                    f, flag = audit_negative(df, amount_col); all_findings.extend(f); flagged_data["Nilai Negatif"] = flag
                    f, flag = audit_outliers_iqr(df, amount_col); all_findings.extend(f); flagged_data["Outlier IQR"] = flag
                if budget_col and amount_col: f, flag = audit_overbudget(df, budget_col, amount_col); all_findings.extend(f); flagged_data["Overbudget"] = flag
                score = min(100, (sum(1 for f in all_findings if f["risk_level"] == "Tinggi") * 20) + (sum(1 for f in all_findings if f["risk_level"] == "Sedang") * 10))
                st.session_state['audit_results'] = {"findings": all_findings, "flagged": flagged_data, "filename": uploaded.name, "entitas": entitas, "periode": periode, "score": score}
                st.session_state.ai_chat_history = []
                flagged_json = {k: v.to_dict(orient="records") for k, v in flagged_data.items() if not v.empty}
                fid = save_file(uploaded.name, uploaded.getvalue())
                aid = save_analysis(fid, "Audit Profesional", ["audit_v5"], {"total_findings": len(all_findings), "score": score, "entitas": entitas, "periode": periode, "flagged_data": flagged_json})
                save_findings(aid, all_findings)
                st.success(f"💾 KKA untuk **{entitas}** berhasil disimpan (ID #{aid}).")
    if 'audit_results' in st.session_state:
        res = st.session_state['audit_results']
        st.markdown("---")
        render_audit_results(res["findings"], res["flagged"], res["filename"], res["entitas"], res["periode"], res["score"])

def page_financial_analysis():
    finora_header("Analisis Laporan Keuangan", "Evaluasi Kesehatan Finansial & Prospek")
    st.markdown("### 1. Upload Laporan Keuangan")
    c1, c2 = st.columns([2, 1])
    uploaded_fin = c1.file_uploader("Upload File Excel/CSV (APBD / Neraca / Laba Rugi)", type=["xlsx", "csv"], key="fin_upload")
    entity_type = c2.selectbox("Jenis Entitas", ["Perusahaan Swasta / Tbk", "Pemerintahan / BLU"])
    if uploaded_fin:
        try:
            df_fin = safe_read_dataframe(uploaded_fin, header_row=0)
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {str(e)}")
            st.stop()
        st.success("✨ File berhasil dibaca.")
        if st.button("📊 ANALISIS KEUANGAN OTOMATIS", use_container_width=True):
            with st.spinner("Menghitung rasio..."):
                nilai_utama = auto_extract_financials(df_fin, entity_type)
                rasio = hitung_rasio_dari_nilai(nilai_utama, entity_type)
                st.session_state['fin_results'] = {"rasio": rasio, "nilai": nilai_utama, "entity_type": entity_type}
    if 'fin_results' in st.session_state:
        res = st.session_state['fin_results']
        rasio, nilai, is_pemda = res["rasio"], res["nilai"], "Pemerintahan" in res['entity_type']
        st.markdown("---")
        tab_rasio, tab_grafik, tab_ai_advisor = st.tabs(["📊 Ringkasan Rasio", "📈 Visualisasi & Prospek", "🧠 AI Financial Advisor"])
        with tab_rasio:
            st.markdown(f"### Kesehatan Finansial: **{res['entity_type']}**")
            if is_pemda:
                c1, c2, c3, c4 = st.columns(4)
                with c1: metric_card("Realisasi Pendapatan", format_rupiah(nilai.get('real_pendapatan', 0)))
                with c2: metric_card("Realisasi Belanja", format_rupiah(nilai.get('real_belanja', 0)))
                with c3: metric_card("Surplus / Defisit", format_rupiah(nilai.get('surplus', 0)))
                with c4: metric_card("Pendapatan Asli Daerah", format_rupiah(nilai.get('pad', 0)))
                st.markdown("#### Rasio Keuangan Sektor Publik")
                r1, r2, r3, r4 = st.columns(4)
                with r1: metric_card("Efektivitas Pendapatan", f"{rasio.get('Efektivitas Pendapatan (%)', 0)} %")
                with r2: metric_card("Efisiensi Belanja", f"{rasio.get('Efisiensi Belanja (%)', 0)} %")
                with r3: metric_card("Kemandirian Daerah", f"{rasio.get('Kemandirian Daerah (%)', 0)} %")
                with r4: metric_card("Pertumbuhan Pendapatan", f"{rasio.get('Pertumbuhan Pendapatan (%)', 0)} %")
                st.markdown("---")
                st.markdown("### 📝 Analisis Kondisi & Rekomendasi Mendalam")
                st.markdown(f"""<div class="insight-box"><h4>1. Efektivitas Pendapatan ({rasio.get('Efektivitas Pendapatan (%)', 0)}%)</h4><p>{get_detailed_insight('Efektivitas', rasio.get('Efektivitas Pendapatan (%)', 0))}</p></div>""", unsafe_allow_html=True)
                st.markdown(f"""<div class="insight-box"><h4>2. Efisiensi Belanja ({rasio.get('Efisiensi Belanja (%)', 0)}%)</h4><p>{get_detailed_insight('Efisiensi', rasio.get('Efisiensi Belanja (%)', 0))}</p></div>""", unsafe_allow_html=True)
                st.markdown(f"""<div class="insight-box"><h4>3. Kemandirian Daerah ({rasio.get('Kemandirian Daerah (%)', 0)}%)</h4><p>{get_detailed_insight('Kemandirian', rasio.get('Kemandirian Daerah (%)', 0))}</p></div>""", unsafe_allow_html=True)
                st.markdown(f"""<div class="insight-box"><h4>4. Pertumbuhan Pendapatan ({rasio.get('Pertumbuhan Pendapatan (%)', 0)}%)</h4><p>{get_detailed_insight('Pertumbuhan', rasio.get('Pertumbuhan Pendapatan (%)', 0))}</p></div>""", unsafe_allow_html=True)
            else:
                c1, c2, c3, c4 = st.columns(4)
                with c1: metric_card("Total Pendapatan", format_rupiah(nilai.get('pendapatan', 0)))
                with c2: metric_card("Laba Bersih", format_rupiah(nilai.get('laba', 0)))
                with c3: metric_card("Total Aset", format_rupiah(nilai.get('total_aset', 0)))
                with c4: metric_card("Total Ekuitas", format_rupiah(nilai.get('ekuitas', 0)))
                st.markdown("#### Rasio Keuangan Utama")
                r1, r2, r3, r4, r5 = st.columns(5)
                with r1: metric_card("Current Ratio", f"{rasio.get('Current Ratio (x)', 0)} x")
                with r2: metric_card("Net Profit Margin", f"{rasio.get('Net Profit Margin (%)', 0)} %")
                with r3: metric_card("ROA", f"{rasio.get('ROA (%)', 0)} %")
                with r4: metric_card("ROE", f"{rasio.get('ROE (%)', 0)} %")
                with r5: metric_card("Debt to Equity", f"{rasio.get('Debt to Equity (x)', 0)} x")
                st.markdown("---")
                st.markdown("### 📝 Analisis Kondisi & Rekomendasi Mendalam")
                st.markdown(f"""<div class="insight-box"><h4>1. Likuiditas: Current Ratio ({rasio.get('Current Ratio (x)', 0)}x)</h4><p>{get_detailed_insight('Current Ratio', rasio.get('Current Ratio (x)', 0))}</p></div>""", unsafe_allow_html=True)
                st.markdown(f"""<div class="insight-box"><h4>2. Solvabilitas: Debt to Equity ({rasio.get('Debt to Equity (x)', 0)}x)</h4><p>{get_detailed_insight('Debt to Equity', rasio.get('Debt to Equity (x)', 0))}</p></div>""", unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("### 📥 Download Laporan Analisis")
            col_ex, col_pdf = st.columns(2)
            col_ex.download_button(label="⬇️ Download Excel (Data & Rasio)", data=generate_fin_excel(nilai, rasio, res['entity_type']), file_name=f"Analisis_Keuangan_{res['entity_type'].replace(' ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            col_pdf.download_button(label="⬇️ Download PDF (Laporan Lengkap)", data=generate_fin_pdf(nilai, rasio, res['entity_type']), file_name=f"Laporan_Analisis_{res['entity_type'].replace(' ', '_')}.pdf", mime="application/pdf", use_container_width=True)
        with tab_grafik:
            st.markdown("### Analisis Visual Prospek")
            if is_pemda:
                fig = go.Figure(data=[go.Bar(name='Anggaran Pendapatan', x=['Pendapatan'], y=[nilai.get('ang_pendapatan', 0)], marker_color='#F8C8D4'), go.Bar(name='Realisasi Pendapatan', x=['Pendapatan'], y=[nilai.get('real_pendapatan', 0)], marker_color='#FF5C8B'), go.Bar(name='Anggaran Belanja', x=['Belanja'], y=[nilai.get('ang_belanja', 0)], marker_color='#F8C8D4'), go.Bar(name='Realisasi Belanja', x=['Belanja'], y=[nilai.get('real_belanja', 0)], marker_color='#7A1F3D')])
                fig.update_layout(barmode='group', title="Anggaran vs Realisasi APBD", height=400)
                st.plotly_chart(fig, use_container_width=True)
            else:
                fig = go.Figure(data=[go.Bar(name='Pendapatan', x=['Keuangan'], y=[nilai.get('pendapatan', 0)], marker_color='#FF5C8B'), go.Bar(name='Laba Bersih', x=['Keuangan'], y=[nilai.get('laba', 0)], marker_color='#7A1F3D')])
                fig.update_layout(barmode='group', title="Pendapatan vs Laba Bersih", height=400)
                st.plotly_chart(fig, use_container_width=True)
        with tab_ai_advisor:
            st.markdown("### 🤖 AI Financial Advisor")
            if st.button("Buat Laporan Analisis AI", use_container_width=True):
                if is_pemda:
                    context_ai = f"Entitas: Pemerintahan / BLU\nRealisasi Pendapatan: Rp {nilai.get('real_pendapatan', 0):,.0f}, Anggaran: Rp {nilai.get('ang_pendapatan', 0):,.0f}\nRealisasi Belanja: Rp {nilai.get('real_belanja', 0):,.0f}, Anggaran: Rp {nilai.get('ang_belanja', 0):,.0f}\nSurplus/Defisit: Rp {nilai.get('surplus', 0):,.0f}\nPAD: Rp {nilai.get('pad', 0):,.0f}\nRasio Efektivitas: {rasio.get('Efektivitas Pendapatan (%)', 0)}%, Efisiensi: {rasio.get('Efisiensi Belanja (%)', 0)}%, Kemandirian: {rasio.get('Kemandirian Daerah (%)', 0)}%"
                    prompt_ai = "Sebagai Konsultan Keuangan Daerah, berikan analisis atas APBD ini (Kesehatan Fiskal, Kemandirian, dan Rekomendasi Optimalisasi Pendapatan/Belanja)."
                else:
                    context_ai = f"Entitas: Perusahaan\nPendapatan: Rp {nilai.get('pendapatan', 0):,.0f}, Laba: Rp {nilai.get('laba', 0):,.0f}\nTotal Aset: Rp {nilai.get('total_aset', 0):,.0f}, Ekuitas: Rp {nilai.get('ekuitas', 0):,.0f}\nRasio: Current Ratio {rasio.get('Current Ratio (x)', 0)}x, NPM {rasio.get('Net Profit Margin (%)', 0)}%, ROA {rasio.get('ROA (%)', 0)}%, ROE {rasio.get('ROE (%)', 0)}%"
                    prompt_ai = "Sebagai Analis Keuangan Senior, berikan analisis kesehatan finansial perusahaan dan rekomendasi strategis ke depan."
                with st.chat_message("assistant"): st.write_stream(ask_gemini_stream(prompt_ai, context_ai, system_role="Konsultan Keuangan & Fiskal Ahli"))
                st.success("Analisis AI Selesai.")

def page_cash_flow():
    finora_header("Cash Flow Intelligence", "Prediksi Likuiditas & Analisis Arus Kas Mendatang")
    
    st.markdown("### 1. Upload Data Transaksi / Laporan")
    c1, c2 = st.columns([3, 1])
    uploaded_cf = c1.file_uploader("Upload File Excel/CSV (Mutasi Bank ATAU Laporan Arus Kas)", type=["xlsx", "csv"], key="cf_upload")
    header_row_cf = c2.number_input("Header di baris ke-", min_value=0, value=0, key="cf_header", help="Ubah ke 1 atau 2 jika baris pertama Excel adalah Judul.")

    def _detect_header_row(df, keywords, max_scan=30):
        for i in range(min(max_scan, len(df))):
            row_vals = [str(c).lower() for c in df.iloc[i].values if pd.notna(c)]
            if any(kw in cell for cell in row_vals for kw in keywords):
                return i
        return None

    if uploaded_cf:
        try:
            # === MUTLAK: BACA EXCEL VIA OPENPYXL (BUKAN PANDAS) ===
            df_raw = read_raw_data(uploaded_cf)

            # === 1. DETEKSI LAPORAN ARUS KAS (STATEMENT) ===
            statement_header_idx = _detect_header_row(
                df_raw, ["uraian", "keterangan", "deskripsi", "pos ", "akun"]
            )
            is_statement = statement_header_idx is not None

            if is_statement:
                st.info("📄 **Terdeteksi: Laporan Arus Kas Komparatif.** Aplikasi akan menampilkan ringkasan laporan.")
                st.session_state.pop('cf_results', None)  # bersihkan hasil lama

                clean_headers = _clean_headers(df_raw.iloc[statement_header_idx].tolist())
                df_cf = df_raw.iloc[statement_header_idx + 1:].copy()
                df_cf.columns = clean_headers
                df_cf = df_cf.reset_index(drop=True)

                col_uraian = next(
                    (c for c in df_cf.columns if any(
                        k in str(c).lower() for k in ["uraian", "keterangan", "deskripsi", "pos"])
                     ), None)
                if col_uraian is None:
                    col_uraian = df_cf.columns[1] if len(df_cf.columns) > 1 else df_cf.columns[0]

                col_2024 = next((c for c in df_cf.columns if "2024" in str(c) or "tahun 1" in str(c).lower()), None)
                col_2023 = next((c for c in df_cf.columns if "2023" in str(c) or "tahun 2" in str(c).lower()), None)
                if not col_2024 and len(df_cf.columns) >= 3:
                    col_2024 = df_cf.columns[2]
                if not col_2023 and len(df_cf.columns) >= 4:
                    col_2023 = df_cf.columns[3]

                if col_2024: df_cf[col_2024] = df_cf[col_2024].apply(clean_currency)
                if col_2023: df_cf[col_2023] = df_cf[col_2023].apply(clean_currency)

                df_cf = df_cf.dropna(subset=[col_uraian])
                df_cf = df_cf[df_cf[col_uraian].astype(str).str.strip() != ""]

                st.write("### Ringkasan Laporan")
                preview_cols = [col_uraian]
                if col_2024: preview_cols.append(col_2024)
                if col_2023: preview_cols.append(col_2023)
                st.dataframe(df_cf[preview_cols].reset_index(drop=True), use_container_width=True)

                main_rows = df_cf[df_cf[col_uraian].astype(str).str.contains(
                    "Arus Kas Bersih|Kenaikan|Penurunan|Saldo Kas", case=False, na=False)]

                if not main_rows.empty:
                    fig = go.Figure()
                    if col_2024: fig.add_trace(go.Bar(name='Tahun 2024', x=main_rows[col_uraian], y=main_rows[col_2024], marker_color='#FF5C8B'))
                    if col_2023: fig.add_trace(go.Bar(name='Tahun 2023', x=main_rows[col_uraian], y=main_rows[col_2023], marker_color='#7A1F3D'))
                    fig.update_layout(barmode='group', title="Perbandingan Arus Kas Utama", height=500)
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("### 🤖 AI Cash Flow Advisor")
                if st.button("Analisis Laporan Arus Kas dengan AI", use_container_width=True):
                    context = "Data Laporan Arus Kas:\n"
                    for _, row in main_rows.iterrows():
                        v24 = row[col_2024] if col_2024 else 0
                        v23 = row[col_2023] if col_2023 else 0
                        context += f"- {row[col_uraian]}: 2024 = Rp {v24:,.0f}, 2023 = Rp {v23:,.0f}\n"
                    prompt = ("Berikan analisis profesional mengenai kesehatan arus kas perusahaan "
                              "berdasarkan data komparatif ini. Soroti tren operasi, investasi, "
                              "pendanaan, dan berikan rekomendasi strategis.")
                    with st.chat_message("assistant"):
                        st.write_stream(ask_gemini_stream(prompt, context, system_role="Analis Arus Kas Ahli"))
                return  # stop di sini untuk mode statement

            # === 2. MODE MUTASI BANK (TRANSAKSI HARIAN) ===
            bank_header_idx = _detect_header_row(
                df_raw, ["tanggal", "tgl", "date", "waktu", "periode"]
            )
            if bank_header_idx is None:
                bank_header_idx = header_row_cf

            clean_headers = _clean_headers(df_raw.iloc[bank_header_idx].tolist())
            df_cf = df_raw.iloc[bank_header_idx + 1:].copy()
            df_cf.columns = clean_headers
            df_cf = df_cf.reset_index(drop=True)

            auto_tgl, auto_ket, auto_nom = "", "", ""
            for c in df_cf.columns:
                c_low = str(c).lower()
                if not auto_tgl and any(k in c_low for k in ["tanggal", "tgl", "date", "waktu", "periode"]): auto_tgl = c
                if not auto_ket and any(k in c_low for k in ["keterangan", "deskripsi", "uraian", "nama", "category", "kategori", "transaksi", "ket"]): auto_ket = c
                if not auto_nom and any(k in c_low for k in ["nominal", "jumlah", "nilai", "amount", "debit", "kredit", "mutasi", "total"]): auto_nom = c

            for c in df_cf.columns:
                sample = df_cf[c].dropna().head(20).astype(str)
                if sample.empty: continue
                if not auto_tgl and pd.to_datetime(sample, errors='coerce', dayfirst=True).notna().sum() > len(sample) * 0.5:
                    auto_tgl = c
                elif not auto_nom and c != auto_tgl:
                    parsed_nums = sample.apply(clean_currency)
                    if (parsed_nums != 0).sum() > len(sample) * 0.5:
                        auto_nom = c

            if not auto_tgl and len(df_cf.columns) > 0: auto_tgl = df_cf.columns[0]
            if not auto_nom and len(df_cf.columns) > 1: auto_nom = df_cf.columns[1]
            if not auto_ket:
                for c in df_cf.columns:
                    if c not in [auto_tgl, auto_nom]:
                        auto_ket = c
                        break

            st.markdown("### 2. Pemetaan Kolom (Otomatis)")
            st.success("✨ FINORA AI telah mendeteksi kolom secara otomatis. Kamu bisa langsung menjalankan analisis.")

            with st.expander("⚙️ Lihat / Koreksi Pemetaan Kolom", expanded=False):
                c3, c4, c5 = st.columns(3)
                col_tgl = c3.selectbox("Kolom Tanggal", list(df_cf.columns),
                                       index=list(df_cf.columns).index(auto_tgl) if auto_tgl in list(df_cf.columns) else 0)
                col_ket = c4.selectbox("Kolom Keterangan", [""] + list(df_cf.columns),
                                       index=list(df_cf.columns).index(auto_ket) + 1 if auto_ket in list(df_cf.columns) else 0)
                col_nom = c5.selectbox("Kolom Nominal", list(df_cf.columns),
                                       index=list(df_cf.columns).index(auto_nom) if auto_nom in list(df_cf.columns) else 0)

            st.markdown("---")
            if st.button("📊 JALANKAN ANALISIS ARUS KAS", use_container_width=True):
                with st.spinner("Memproses data..."):
                    df_cf['_parsed_date'] = pd.to_datetime(df_cf[col_tgl], errors='coerce', dayfirst=True)
                    if df_cf['_parsed_date'].isna().all() and pd.api.types.is_numeric_dtype(df_cf[col_tgl]):
                        df_cf['_parsed_date'] = pd.to_datetime(df_cf[col_tgl], errors='coerce', unit='D', origin='1899-12-30')

                    df_cf[col_nom] = df_cf[col_nom].apply(clean_currency)

                    if df_cf['_parsed_date'].isna().all():
                        st.error(f"❌ Gagal membaca kolom tanggal **{col_tgl}**. Pastikan kolom tersebut berisi tanggal. Silakan ganti pilihan kolom di menu 'Koreksi Pemetaan Kolom' di atas.")
                    else:
                        df_cf = df_cf.dropna(subset=['_parsed_date']).sort_values(by='_parsed_date')
                        if col_ket:
                            df_cf['Kategori'] = df_cf[col_ket].astype(str).str.strip()
                        else:
                            df_cf['Kategori'] = 'Lainnya'

                        total_masuk = df_cf[df_cf[col_nom] > 0][col_nom].sum()
                        total_keluar = df_cf[df_cf[col_nom] < 0][col_nom].sum()
                        df_cf['Running Balance'] = df_cf[col_nom].cumsum()
                        saldo_akhir = df_cf['Running Balance'].iloc[-1] if not df_cf.empty else 0

                        bulan_unik = max(1, df_cf['_parsed_date'].dt.to_period('M').nunique())
                        burn_rate = abs(total_keluar) / bulan_unik
                        bulan_bertahan = (saldo_akhir / burn_rate) if burn_rate > 0 and saldo_akhir > 0 else 999

                        st.session_state['cf_results'] = {
                            "df": df_cf, "col_tgl": '_parsed_date', "col_nom": col_nom,
                            "total_masuk": total_masuk, "total_keluar": total_keluar,
                            "saldo_akhir": saldo_akhir, "burn_rate": burn_rate,
                            "bulan_bertahan": bulan_bertahan
                        }
                        st.success("✅ Analisis berhasil dijalankan!")
                        st.rerun()

            with st.expander("👀 Preview Data Mentah (Buka jika hasil 0)", expanded=False):
                st.write(f"Preview kolom **{col_tgl}** (Tanggal) dan **{col_nom}** (Nominal):")
                preview_cols = [col_tgl, col_nom]
                if col_ket: preview_cols.insert(1, col_ket)
                st.dataframe(df_cf[preview_cols].head(10))
                
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {str(e)}. Pastikan format file sesuai.")

    if 'cf_results' in st.session_state:
        res_cf = st.session_state['cf_results']
        st.markdown("---")
        st.markdown("### 📊 Kartu Likuiditas")
        m1, m2, m3, m4 = st.columns(4)
        with m1: metric_card("Total Masuk", format_rupiah(res_cf['total_masuk']))
        with m2: metric_card("Total Keluar", format_rupiah(res_cf['total_keluar']))
        with m3: metric_card("Saldo Akhir (Net)", format_rupiah(res_cf['saldo_akhir']))
        with m4: metric_card("Cash Runway", f"{res_cf['bulan_bertahan']:.1f} Bulan", help_text="Estimasi berapa bulan kas Anda akan bertahan.")
        
        tab_grafik_cf, tab_deteksi_cf, tab_ai_cf = st.tabs(["📈 Tren & Kategori", "🚨 Deteksi Anomali", "🧠 AI Cash Flow Advisor"])
        
        with tab_grafik_cf:
            st.markdown("#### Tren Saldo Kas (Net)")
            fig_area = px.area(res_cf['df'], x=res_cf['col_tgl'], y='Running Balance', title="Pergerakan Saldo Kas (Akumulasi Net)", color_discrete_sequence=['#FF5C8B'])
            fig_area.update_traces(fillcolor='rgba(255, 92, 139, 0.2)', line_color='#7A1F3D', line_width=3)
            st.plotly_chart(fig_area, use_container_width=True)
            
            st.markdown("#### Proporsi Pengeluaran (Top 10 Kategori)")
            df_out = res_cf['df'][res_cf['df'][res_cf['col_nom']] < 0].copy()
            df_out['Nominal_Abs'] = df_out[res_cf['col_nom']].abs()
            top_out = df_out.groupby('Kategori')['Nominal_Abs'].sum().nlargest(10).reset_index()
            if not top_out.empty: st.plotly_chart(px.pie(top_out, names='Kategori', values='Nominal_Abs', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu_r), use_container_width=True)
            else: st.info("Tidak ada data pengeluaran.")
            
        with tab_deteksi_cf:
            st.markdown("### 🚨 Transaksi Anomali (Pengeluaran Ekstrem)")
            df_out_all = res_cf['df'][res_cf['df'][res_cf['col_nom']] < 0]
            if not df_out_all.empty:
                q1 = df_out_all[res_cf['col_nom']].quantile(0.25); q3 = df_out_all[res_cf['col_nom']].quantile(0.75); iqr = q3 - q1
                df_anomali = df_out_all[df_out_all[res_cf['col_nom']] < (q1 - 1.5 * iqr)]
                if not df_anomali.empty: st.dataframe(df_anomali[[res_cf['col_tgl'], 'Kategori', res_cf['col_nom']]], use_container_width=True)
                else: st.success("Aman! Tidak ada pengeluaran ekstrem yang terdeteksi.")
            else: st.info("Tidak ada data pengeluaran untuk dianalisis.")
            
        with tab_ai_cf:
            st.markdown("### 🤖 AI Cash Flow Advisor")
            if st.button("Analisis Kesehatan Cash Flow", use_container_width=True):
                context_cf = f"Data Arus Kas:\n- Total Pemasukan: Rp {res_cf['total_masuk']:,.0f}\n- Total Pengeluaran: Rp {res_cf['total_keluar']:,.0f}\n- Saldo Akhir (Net): Rp {res_cf['saldo_akhir']:,.0f}\n- Rata-rata Pengeluaran Bulanan: Rp {res_cf['burn_rate']:,.0f}\n- Estimasi Kas Bertahan (Runway): {res_cf['bulan_bertahan']:.1f} Bulan"
                prompt_cf = "Sebagai Konsultan Keuangan, berikan analisis atas arus kas ini.\n1. Berapa lama perusahaan bisa bertahan berdasarkan arus kas bersihnya?\n2. Apa saran konkret untuk memperpanjang cash runway?\n3. Berikan simulasi: 'Jika memotong biaya operasional 20%, berapa lama cash flow bertahan?'"
                with st.chat_message("assistant"): st.write_stream(ask_gemini_stream(prompt_cf, context_cf, system_role="Konsultan Keuangan & Cash Flow Specialist"))
                st.success("Analisis AI Selesai.")

def page_learning_center():
    finora_header("Akademi FINORA", "Pusat Pembelajaran Audit & Keuangan Perusahaan")
    st.markdown("Selamat datang di **Akademi FINORA**! Di sini Anda dapat mempelajari dasar-dasar audit dan analisis keuangan perusahaan.")
    st.markdown("---")
    tab_audit, tab_keuangan = st.tabs(["🔎 Modul Audit", "🏢 Modul Keuangan Perusahaan"])
    with tab_audit:
        st.markdown("### 📚 Kurikulum Dasar Audit")
        with st.expander("Modul 1: Apa itu Audit dan Mengapa Penting?"):
            st.markdown("**Pengertian Audit:**\nAudit adalah proses pengumpulan dan evaluasi bukti secara objektif mengenai asersi tentang tindakan dan peristiwa ekonomi, untuk menentukan tingkat kesesuaian antara asersi tersebut dengan kriteria yang telah ditetapkan.\n\n**Mengapa Penting?**\n1. **Mengurangi Risiko Informasi:** Manajemen mungkin memiliki insentif untuk memanipulasi laporan. Auditor independen memvalidasi kebenarannya.\n2. **Mendeteksi Kecurangan (Fraud):** Menemukan anomali seperti penggelapan kas, mark-up harga, atau transaksi fiktif.\n3. **Kepatuhan (Compliance):** Memastikan perusahaan mematuhi peraturan pemerintah dan standar akuntansi (SAK/PSAK).")
        with st.expander("Modul 2: 5 Jenis Opini Audit"):
            st.markdown("Auditor memberikan opini atas kewajaran laporan keuangan. Ada 5 jenis opini:\n1. **Wajar Tanpa Modifikasian (Unqualified Opinion):** Laporan keuangan disajikan secara wajar. (Opini terbaik).\n2. **Wajar Dengan Pengecualian (Qualified Opinion):** Ada masalah material, tapi tidak mempengaruhi keseluruhan laporan. (Misal: Salah satu cabang tidak diaudit).\n3. **Opini Tanpa Wajar (Adverse Opinion):** Laporan keuangan salah saji secara material dan merata. (Sangat buruk).\n4. **Opini Tidak Menyatakan Pendapat (Disclaimer of Opinion):** Auditor tidak bisa mendapatkan bukti yang cukup untuk memberikan opini.\n5. **Opini Wajar Tanpa Modifikasian dengan Paragraf Penekanan:** Ada hal penting yang perlu ditekankan (misal: bencana alam).")
        with st.expander("Modul 3: Teknik Sampling dalam Audit"):
            st.markdown("Auditor tidak mungkin memeriksa seluruh transaksi jika jumlahnya jutaan. Oleh karena itu, digunakan **Teknik Sampling**.\n*   **Statistical Sampling:** Menggunakan rumus matematika (seperti *Random Sampling* atau *Monetary Unit Sampling*). Cocok untuk data yang sangat besar dan homogen.\n*   **Non-Statistical Sampling (Judgmental):** Berdasarkan pengalaman auditor. (Misal: memilih transaksi di atas Rp 100 juta karena risiko lebih tinggi).\n*   **Interval Sampling:** Memilih sampel berdasarkan kelipatan angka tertentu (misal: setiap transaksi ke-50).")
    with tab_keuangan:
        st.markdown("### 📈 Kurikulum Analisis Keuangan")
        with st.expander("Modul 1: 3 Laporan Keuangan Utama"):
            st.markdown("Untuk menganalisis perusahaan, Anda harus paham 3 laporan inti:\n1. **Laporan Laba Rugi (Income Statement):** Menunjukkan kinerja operasional selama satu periode. (Pendapatan - Beban = Laba Bersih).\n2. **Neraca (Balance Sheet):** Potret posisi keuangan pada satu titik waktu tertentu. (Aset = Liabilitas + Ekuitas).\n3. **Laporan Arus Kas (Cash Flow Statement):** Melacak aliran kas masuk dan keluar (Operasi, Investasi, Pendanaan). *Ingat: Laba tinggi belum tentu kasnya ada!*")
        with st.expander("Modul 2: Memahami Rasio Keuangan (Financial Ratios)"):
            st.markdown("Rasio keuangan membantu kita menilai kesehatan perusahaan tanpa harus membaca seluruh angka. Terdapat 4 kategori utama:\n*   **Rasio Likuiditas (Current Ratio):** Apakah perusahaan mampu membayar utang jangka pendek? (Aset Lancar / Utang Lancar).\n*   **Rasio Solvabilitas (Debt to Equity):** Seberapa besar perusahaan bergantung pada utang? (Total Utang / Ekuitas).\n*   **Rasio Profitabilitas (ROA, ROE, NPM):** Seberapa efisien perusahaan menghasilkan laba? (Laba / Aset atau Ekuitas).\n*   **Rasio Aktivitas (Perputaran Piutang):** Seberapa cepat perusahaan menagih piutangnya?")
        with st.expander("Modul 3: Deteksi Dini Kebangkrutan"):
            st.markdown("Perusahaan bisa bangkrut bukan hanya karena rugi, tapi seringkali karena **Kehabisan Kas (Cash Flow Trap)**.\n\n**Tanda-tanda Bahaya:**\n*   Laba di laporan naik, tapi saldo kas menurun drastis.\n*   Piutang menumpuk tak tertagih (banyak penjualan fiktif).\n*   Ketergantungan tinggi pada utang jangka pendek untuk membiayai operasional.\n\n**Solusi:**\nGunakan analisis *Cash Runway* (seperti yang ada di menu Analisis Arus Kas FINORA) untuk melihat seberapa lama kas perusahaan bisa bertahan sebelum habis.")

def page_dashboard():
    finora_header("Dashboard", "Ringkasan seluruh aktivitas FINORA AI")
    conn = get_conn()
    t_files = pd.read_sql_query("SELECT COUNT(*) c FROM files", conn).iloc[0]['c']
    t_analyses = pd.read_sql_query("SELECT COUNT(*) c FROM analyses", conn).iloc[0]['c']
    t_findings = pd.read_sql_query("SELECT COUNT(*) c FROM findings", conn).iloc[0]['c']
    h_risk = pd.read_sql_query("SELECT COUNT(*) c FROM findings WHERE risk_level='Tinggi'", conn).iloc[0]['c']
    conn.close()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📁 File Dianalisis", t_files); c2.metric("📊 Analisis Dilakukan", t_analyses); c3.metric("⚠️ Total Temuan", t_findings); c4.metric("🔴 Risiko Tinggi", h_risk)
    st.markdown("---")
    st.subheader("🕐 Analisis Terakhir")
    recent = fetch_analyses()
    if not recent.empty:
        recent['created_at'] = pd.to_datetime(recent['created_at']).dt.strftime("%Y-%m-%d %H:%M")
        st.dataframe(recent.head(5), use_container_width=True)

def page_simpan():
    finora_header("Pusat Laporan & Arsip", "Riwayat analisis dan Kertas Kerja Audit tersimpan")
    df = fetch_analyses()
    if df.empty:
        st.info("Belum ada riwayat."); return
    st.dataframe(df, use_container_width=True)
    st.markdown("---")
    aid = st.selectbox("Pilih ID Analisis untuk membuka KKA lengkap", df["id"].tolist())
    if aid:
        an = fetch_analysis_by_id(aid)
        if an:
            summary_data = json.loads(an["summary_json"])
            entitas, periode, score = summary_data.get("entitas", "-"), summary_data.get("periode", "-"), summary_data.get("score", 0)
            flagged_data = {k: pd.DataFrame(v) for k, v in summary_data.get("flagged_data", {}).items()}
            findings_df = fetch_findings(aid)
            findings_list = findings_df.to_dict("records")
            render_audit_results(findings_list, flagged_data, an["filename"], entitas, periode, score)

def page_landing():
    st.markdown("""<style>[data-testid="collapsedControl"], [data-testid="stSidebar"] { display: none; }</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists(LOGO_FILENAME): st.image(LOGO_FILENAME, width=250)
        else: st.markdown("<h1 style='text-align:center; font-size: 5rem;'>F</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #FFF0F4 0%, #FFF5E8 100%); border-radius: 20px; margin-bottom: 40px; border: 2px solid #FF5C8B; box-shadow: 0 10px 30px rgba(255, 92, 139, 0.15);">
        <h1 style="color: #7A1F3D; font-size: 3.5rem; margin-bottom: 15px; font-weight: 800;">FINORA AI</h1>
        <p style="color: #2D1B24; font-size: 1.3rem; max-width: 800px; margin: 0 auto; line-height: 1.6;">
            Platform Kecerdasan Finansial untuk mendeteksi anomali (Audit) dan menyusun <b>Kertas Kerja Audit (KKA)</b> serta <b>Analisis Laporan Keuangan</b> secara otomatis.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #7A1F3D; margin-bottom: 30px; font-weight: 700;'>Dua Senjata Utama FINORA</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""<div style="background: white; padding: 30px 20px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); text-align: center; border-top: 5px solid #FF5C8B;"><div style="font-size: 3rem; margin-bottom: 15px;">🔎</div><h3 style="color: #7A1F3D; margin-bottom: 10px;">Audit Profesional</h3><p style="color: #555; font-size: 0.95rem;">Mendeteksi duplikasi, transaksi weekend, overbudget, outlier, dan anomali lainnya. Hasil siap diunduh dalam format KKA (PDF & Excel).</p></div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""<div style="background: white; padding: 30px 20px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); text-align: center; border-top: 5px solid #FF5C8B;"><div style="font-size: 3rem; margin-bottom: 15px;">📈</div><h3 style="color: #7A1F3D; margin-bottom: 10px;">Analisis Keuangan</h3><p style="color: #555; font-size: 0.95rem;">Hitung rasio (Likuiditas, Profitabilitas, Solvabilitas) dan dapatkan *insight* prospek bisnis/pemerintahan dari AI Financial Advisor.</p></div>""", unsafe_allow_html=True)
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    col_cta1, col_cta2, col_cta3 = st.columns([1, 2, 1])
    with col_cta2:
        st.markdown("""<style>.stButton > button { background: linear-gradient(135deg, #FF5C8B 0%, #7A1F3D 100%) !important; color: white !important; font-size: 1.3rem !important; padding: 20px 30px !important; border-radius: 12px !important; border: none !important; font-weight: 800 !important; box-shadow: 0 6px 20px rgba(255, 92, 139, 0.4) !important; transition: all 0.3s ease !important; width: 100%; }</style>""", unsafe_allow_html=True)
        if st.button("🚀 START YOUR ANALYSIS", use_container_width=True):
            st.session_state.app_started = True
            st.rerun()

def main():
    if "app_started" not in st.session_state:
        st.session_state.app_started = False
    if not st.session_state.app_started:
        page_landing()
    else:
        try:
            st.sidebar.image(LOGO_FILENAME, width=180)
        except:
            st.sidebar.markdown('<h1 style="color:#FF5C8B;text-align:center;">F</h1>', unsafe_allow_html=True)
        st.sidebar.markdown('<p style="text-align:center;font-size:12px;color:#7A1F3D;">Financial Intelligence<br>Platform</p>', unsafe_allow_html=True)
        st.sidebar.markdown("---")
        nav = st.sidebar.radio("Navigasi", [
            "Dashboard", 
            "Audit Profesional", 
            "Analisis Keuangan", 
            "Analisis Arus Kas", 
            "Akademi FINORA", 
            "Simpan"
        ])
        st.sidebar.markdown("---")
        if st.sidebar.button("🏠 Kembali ke Halaman Depan", use_container_width=True):
            st.session_state.app_started = False
            st.rerun()
        st.sidebar.caption("⚠️ Temuan bersifat indikasi dan perlu verifikasi dokumen fisik.")
        if nav == "Dashboard": page_dashboard()
        elif nav == "Audit Profesional": page_audit()
        elif nav == "Analisis Keuangan": page_financial_analysis()
        elif nav == "Analisis Arus Kas": page_cash_flow()
        elif nav == "Akademi FINORA": page_learning_center()
        elif nav == "Simpan": page_simpan()

if __name__ == "__main__":
    main()
