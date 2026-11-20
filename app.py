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
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "-"
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
    if isinstance(x, (int, float, np.number)): return float(x)
    x = str(x).strip()
    if not x or x.lower() in ("nan", "none", "-", "–", "—"):
        return 0.0
    is_negative = False
    if x.startswith('(') and x.endswith(')'):
        is_negative = True
        x = x[1:-1].strip()
    x = re.sub(r'[^\d,.-]', '', x)
    if not x: return 0.0
    if ',' in x and '.' in x:
        if x.rfind('.') > x.rfind(','): x = x.replace(',', '')
        else: x = x.replace('.', '').replace(',', '.')
    elif ',' in x:
        if len(x.split(',')[-1]) == 3 and x.count(',') >= 1 and not x.startswith('0,'):
            x = x.replace(',', '')
        else:
            x = x.replace(',', '.')
    elif '.' in x:
        if len(x.split('.')[-1]) == 3 and x.count('.') > 0 and not x.startswith('0.'):
            x = x.replace('.', '')
    try:
        val = float(x)
        return -val if is_negative else val
    except Exception:
        return 0.0

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
def read_excel_sheets(uploaded_file):
    """Membaca semua sheet dari file Excel/CSV dan mengembalikan dict {sheet_name: DataFrame}."""
    uploaded_file.seek(0)
    bytes_data = uploaded_file.read()
    fname = uploaded_file.name.lower()

    if fname.endswith(('.xlsx', '.xlsm')):
        file_stream = io.BytesIO(bytes_data)
        wb = openpyxl.load_workbook(file_stream, data_only=True, read_only=True)
        sheets = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = list(ws.iter_rows(values_only=True))
            if data:
                df_sheet = pd.DataFrame(data).dropna(how="all").reset_index(drop=True)
                if not df_sheet.empty:
                    sheets[sheet_name] = df_sheet
        wb.close()
        return sheets if sheets else {"Sheet1": pd.DataFrame()}
    else:
        for sep in [',', ';', '\t']:
            try:
                file_stream = io.BytesIO(bytes_data)
                df_csv = pd.read_csv(file_stream, header=None, sep=sep, on_bad_lines='skip')
                if df_csv.shape[1] > 1 or sep == '\t':
                    return {"CSV": df_csv.dropna(how="all").reset_index(drop=True)}
            except Exception:
                continue
        file_stream = io.BytesIO(bytes_data)
        return {"CSV": pd.read_csv(file_stream, header=None, on_bad_lines='skip').dropna(how="all").reset_index(drop=True)}

def read_raw_data(uploaded_file, sheet_name=None):
    """Membaca file mentah tanpa memicu pandas duplicate column error."""
    sheets = read_excel_sheets(uploaded_file)
    if sheet_name and sheet_name in sheets:
        return sheets[sheet_name]
    best_df = pd.DataFrame()
    max_rows = -1
    for df_s in sheets.values():
        if len(df_s) > max_rows:
            max_rows = len(df_s)
            best_df = df_s
    return best_df

def safe_read_dataframe(uploaded, header_row=0, sheet_name=None):
    df_raw = read_raw_data(uploaded, sheet_name=sheet_name)
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
    """Analisis arus kas berbasis laporan arus kas (PSAK 207) atau transaksi kas."""
    finora_header(
        "Cash Flow Intelligence",
        "Analisis arus kas operasi, investasi, pendanaan, rasio, tren, risiko, dan insight AI"
    )

    st.info(
        "Format laporan mengacu pada struktur PSAK 207 (sebelumnya PSAK 2): "
        "aktivitas operasi, investasi, dan pendanaan. Hasil otomatis tetap perlu "
        "ditinjau karena klasifikasi transaksi bergantung pada karakter transaksi dan kebijakan entitas."
    )

    uploaded_cf = st.file_uploader(
        "Upload Laporan Arus Kas atau Data Transaksi (Excel/CSV)",
        type=["xlsx", "csv"],
        key="cf_upload_v2"
    )
    if not uploaded_cf:
        st.markdown("#### Format data yang dapat digunakan")
        st.markdown(
            "- **Laporan arus kas:** kolom uraian/keterangan dan satu atau lebih kolom nominal/periode.\n"
            "- **Data transaksi:** tanggal, keterangan, kas masuk dan kas keluar atau satu kolom nominal bertanda positif/negatif.\n"
            "- **Laporan komparatif:** dapat memuat dua periode atau lebih untuk analisis tren."
        )
        return

    try:
        all_sheets = read_excel_sheets(uploaded_cf)
        sheet_names = list(all_sheets.keys())
        if len(sheet_names) > 1:
            default_idx = max(range(len(sheet_names)), key=lambda i: len(all_sheets[sheet_names[i]]))
            chosen_sheet = st.selectbox("Pilih Sheet Excel", sheet_names, index=default_idx, key="cf_sheet_select")
            raw = all_sheets[chosen_sheet]
        else:
            chosen_sheet = sheet_names[0]
            raw = all_sheets[chosen_sheet]
        if raw.empty:
            st.error("File tidak berisi data yang dapat dibaca.")
            return
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return

    def detect_header(frame, keywords, max_scan=100):
        best_idx = 0
        best_score = -1
        for i in range(min(max_scan, len(frame))):
            vals = [str(v).strip().lower() for v in frame.iloc[i].tolist() if pd.notna(v) and str(v).strip() != ""]
            if not vals:
                continue
            score = 0
            for val in vals:
                if any(k in val for k in keywords):
                    score += 3
                if re.search(r"\b(20\d{2}|19\d{2})\b", val) or "tahun" in val or "periode" in val:
                    score += 2
            if score > best_score and score > 0:
                best_score = score
                best_idx = i
        return best_idx

    def to_number(value):
        return clean_currency(value)

    def guess_line_activity(label):
        low = str(label).lower()
        if any(k in low for k in [
            "aset tetap", "perolehan aset", "pembelian aset", "penjualan aset",
            "investasi", "deposito", "properti investasi", "aset takberwujud", "capex"
        ]):
            return "investasi"
        if any(k in low for k in [
            "pinjaman", "utang bank", "obligasi", "modal saham", "setoran modal",
            "dividen", "sewa pembiayaan", "liabilitas sewa", "saham treasuri", "pendanaan"
        ]):
            return "pendanaan"
        if any(k in low for k in [
            "pelanggan", "pemasok", "karyawan", "beban operasi", "pajak penghasilan",
            "bunga dibayar", "bunga diterima", "operasi", "kas dari aktivitas operasi",
            "penerimaan kas", "pembayaran kas", "piutang", "persediaan", "utang usaha"
        ]):
            return "operasi"
        return "other"

    def classify_cash_flow(label, current_section=""):
        s = str(label).strip().lower()
        s_clean = re.sub(r"\s+", " ", s)
        if not s_clean or s_clean in ("nan", "none", "-", "–", "—"):
            return current_section, "skip"

        # Deteksi header bagian PSAK 207
        is_subtotal_phrase = any(k in s_clean for k in [
            "kas yang dihasilkan dari", "kas neto diperoleh dari", "kas neto yang diperoleh dari",
            "kas neto digunakan untuk", "kas neto yang digunakan untuk",
            "jumlah arus kas", "total arus kas", "subtotal"
        ])
        if not is_subtotal_phrase:
            if ("aktivitas operasi" in s_clean or "arus kas operasi" in s_clean) and len(s_clean) <= 75:
                return "operasi", "section"
            if ("aktivitas investasi" in s_clean or "arus kas investasi" in s_clean) and len(s_clean) <= 75:
                return "investasi", "section"
            if ("aktivitas pendanaan" in s_clean or "arus kas pendanaan" in s_clean) and len(s_clean) <= 75:
                return "pendanaan", "section"

        # Deteksi baris rekonsiliasi saldo kas
        if any(k in s_clean for k in [
            "kenaikan neto kas", "penurunan neto kas", "kenaikan (penurunan) neto kas",
            "kenaikan/(penurunan) neto kas", "kenaikan bersih kas", "penurunan bersih kas",
            "perubahan neto kas", "kas dan setara kas pada awal", "kas dan setara kas awal",
            "saldo kas awal", "kas awal tahun", "kas awal periode",
            "kas dan setara kas pada akhir", "kas dan setara kas akhir",
            "saldo kas akhir", "kas akhir tahun", "kas akhir periode",
            "pengaruh perubahan kurs", "selisih kurs"
        ]):
            return current_section, "recon"

        # Lewati baris subtotal bagian agar tidak double-counting
        if is_subtotal_phrase:
            return current_section, "subtotal"

        if current_section in ("operasi", "investasi", "pendanaan"):
            return current_section, current_section

        # Fallback jika laporan tidak memiliki judul bagian eksplisit
        inferred = guess_line_activity(s_clean)
        return current_section, inferred

    def make_cf_excel(statement_df, ratios, summary, trend_df=None, anomalies=None):
        buff = io.BytesIO()
        with pd.ExcelWriter(buff, engine="openpyxl") as writer:
            statement_df.to_excel(writer, sheet_name="Analisis Arus Kas", index=False)
            pd.DataFrame([{"Indikator": k, "Nilai": v} for k, v in summary.items()]).to_excel(
                writer, sheet_name="Ringkasan", index=False
            )
            if ratios:
                pd.DataFrame([{"Rasio": k, "Nilai": v} for k, v in ratios.items()]).to_excel(
                    writer, sheet_name="Rasio Arus Kas", index=False
                )
            if trend_df is not None and not trend_df.empty:
                trend_df.to_excel(writer, sheet_name="Tren", index=False)
            if anomalies is not None and not anomalies.empty:
                anomalies.to_excel(writer, sheet_name="Anomali", index=False)
            for sheet in writer.book.worksheets:
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions
                for col_cells in sheet.columns:
                    max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
                    sheet.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 14), 48)
        buff.seek(0)
        return buff

    def make_cf_pdf(title, summary, detail_df):
        buff = io.BytesIO()
        doc = SimpleDocTemplate(buff, pagesize=A4, rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "CFTitle", parent=styles["Title"], textColor=colors.HexColor("#7A1F3D"),
            alignment=1, fontSize=16, leading=20
        )
        heading_style = ParagraphStyle(
            "CFHeading", parent=styles["Heading2"], textColor=colors.HexColor("#7A1F3D"),
            fontSize=11, leading=14
        )
        elements = [
            Paragraph("FINORA AI", title_style),
            Paragraph(title, heading_style),
            Paragraph(f"Tanggal dibuat: {datetime.now().strftime('%d-%m-%Y %H:%M')}", styles["Normal"]),
            Spacer(1, 12)
        ]
        summary_rows = [["Indikator", "Nilai"]]
        for key, value in summary.items():
            shown = format_rupiah(value) if isinstance(value, (int, float, np.number)) else str(value if value is not None else "-")
            summary_rows.append([str(key), shown])
        summary_table = Table(summary_rows, colWidths=[230, 260], repeatRows=1)
        summary_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7A1F3D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FFF5E8")])
        ]))
        elements.extend([summary_table, Spacer(1, 14), Paragraph("Rincian Data", heading_style)])
        if detail_df is not None and not detail_df.empty:
            preview = detail_df.head(40).copy()
            pdf_rows = [[str(c) for c in preview.columns]]
            for row in preview.itertuples(index=False, name=None):
                pdf_rows.append([str(v)[:80] for v in row])
            detail_table = Table(pdf_rows, repeatRows=1)
            detail_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FF5C8B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP")
            ]))
            elements.append(detail_table)
        else:
            elements.append(Paragraph("Tidak ada rincian tambahan.", styles["Normal"]))
        doc.build(elements)
        buff.seek(0)
        return buff

    mode = st.radio(
        "Jenis data yang diunggah",
        ["Deteksi otomatis", "Laporan arus kas (statement)", "Data transaksi / mutasi bank"],
        horizontal=True,
        key="cf_mode"
    )
    file_sig = f"{uploaded_cf.name}_{chosen_sheet}_{raw.shape}"
    header_guess = detect_header(raw, ["uraian", "keterangan", "deskripsi", "tanggal", "date", "pos", "akun"])
    header_row = st.number_input(
        "Baris header kolom (mulai dari 0)",
        min_value=0,
        max_value=max(0, len(raw) - 1),
        value=int(header_guess),
        key=f"cf_header_{file_sig}"
    )
    headers = _clean_headers(raw.iloc[int(header_row)].tolist())
    data = raw.iloc[int(header_row) + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all").reset_index(drop=True)

    statement_keywords = ["uraian", "keterangan", "deskripsi", "pos", "akun"]
    detected_statement = any(any(k in str(c).lower() for k in statement_keywords) for c in headers)
    use_statement = mode == "Laporan arus kas (statement)" or (
        mode == "Deteksi otomatis" and detected_statement and
        not any(any(k in str(c).lower() for k in ["tanggal", "tgl", "date"]) for c in headers)
    )

    if use_statement:
        st.markdown("### 1. Pemetaan laporan arus kas")
        desc_candidates = [c for c in data.columns if any(k in str(c).lower() for k in ["uraian", "keterangan", "deskripsi", "pos", "akun"])]
        desc_default = desc_candidates[0] if desc_candidates else data.columns[0]
        desc_col = st.selectbox(
            "Kolom uraian / keterangan",
            list(data.columns),
            index=list(data.columns).index(desc_default),
            key=f"cf_stmt_desc_{file_sig}_{header_row}"
        )

        # Pindai SELURUH baris (bukan hanya 30 baris pertama) + deteksi header tahun/periode
        numeric_candidates = []
        year_candidates = []
        for c in data.columns:
            if c == desc_col:
                continue
            col_lower = str(c).lower()
            is_year_header = bool(re.search(r"\b(20\d{2}|19\d{2})\b", col_lower)) or any(
                k in col_lower for k in ["tahun", "periode", "nominal", "jumlah", "nilai", "rp", "idr"]
            )
            vals = data[c].dropna()
            non_zero_count = sum(to_number(v) != 0 for v in vals)
            if non_zero_count >= 1 or is_year_header:
                numeric_candidates.append(c)
                if is_year_header:
                    year_candidates.append(c)

        if not numeric_candidates:
            numeric_candidates = [c for c in data.columns if c != desc_col]
        if not numeric_candidates:
            st.error("Kolom nominal/periode tidak ditemukan. Periksa struktur file.")
            return

        default_periods = year_candidates if year_candidates else numeric_candidates
        selected_periods = st.multiselect(
            "Pilih kolom nominal/periode yang akan dianalisis",
            numeric_candidates,
            default=default_periods,
            key=f"cf_stmt_periods_{file_sig}_{header_row}"
        )

        with st.expander(f"👁️ Pratinjau Seluruh Data Terbaca ({len(data)} baris)", expanded=False):
            st.dataframe(data, use_container_width=True)

        if not selected_periods:
            st.warning("Pilih minimal satu kolom nominal/periode.")
            return

        if st.button("📊 ANALISIS LAPORAN ARUS KAS", use_container_width=True, key="run_cf_statement"):
            statement_results = {}
            for period_col in selected_periods:
                section = ""
                totals = {"operasi": 0.0, "investasi": 0.0, "pendanaan": 0.0}
                opening_cash, closing_cash, reported_net_change, fx_effect = None, None, None, 0.0
                line_items = []
                unclassified_items = []

                for _, row in data.iterrows():
                    raw_label = row[desc_col]
                    if pd.isna(raw_label):
                        continue
                    label = str(raw_label).strip()
                    if not label or label.lower() in ("nan", "none"):
                        continue

                    value = to_number(row[period_col])
                    section, kind = classify_cash_flow(label, section)

                    if kind in ("operasi", "investasi", "pendanaan"):
                        totals[kind] += value
                        line_items.append({"Aktivitas": kind.title(), "Uraian": label, "Nilai": value})
                    elif kind == "recon":
                        low = label.lower()
                        if "awal" in low:
                            opening_cash = value
                        elif "akhir" in low:
                            closing_cash = value
                        elif "kurs" in low:
                            fx_effect = value
                        elif any(k in low for k in ["kenaikan", "penurunan", "perubahan"]):
                            reported_net_change = value
                    elif kind == "other" and value != 0:
                        unclassified_items.append({"Uraian": label, "Nilai": value})

                net_change = sum(totals.values())
                if net_change == 0 and reported_net_change is not None:
                    net_change = reported_net_change
                calculated_closing = (opening_cash + net_change + fx_effect) if opening_cash is not None else None
                if closing_cash is None and calculated_closing is not None:
                    closing_cash = calculated_closing

                capex = abs(sum(
                    x["Nilai"] for x in line_items
                    if x["Aktivitas"] == "Investasi" and any(
                        k in x["Uraian"].lower() for k in [
                            "pembelian aset tetap", "perolehan aset tetap", "pembelian aset",
                            "perolehan aset", "capital expenditure", "capex"
                        ]
                    )
                ))
                fcf = totals["operasi"] - capex
                statement_results[period_col] = {
                    "totals": totals,
                    "opening_cash": opening_cash,
                    "closing_cash": closing_cash,
                    "calculated_closing": calculated_closing,
                    "fx_effect": fx_effect,
                    "net_change": net_change,
                    "fcf": fcf,
                    "line_items": line_items,
                    "unclassified_items": unclassified_items
                }

            st.session_state["cf_statement_results"] = {
                "results": statement_results,
                "filename": uploaded_cf.name,
                "desc_col": desc_col,
                "total_rows_read": len(data)
            }
            st.success(f"Analisis laporan arus kas selesai ({len(statement_results)} periode & {len(data)} baris terbaca).")

        if "cf_statement_results" in st.session_state:
            saved = st.session_state["cf_statement_results"]
            st.markdown("---")
            st.markdown("### 2. Cash Flow Overview")
            periods = list(saved["results"].keys())
            chosen = st.selectbox("Periode yang ditampilkan", periods, key="cf_stmt_selected_period")
            item = saved["results"][chosen]
            totals = item["totals"]
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Kas Neto Operasi", format_rupiah(totals["operasi"]))
            k2.metric("Kas Neto Investasi", format_rupiah(totals["investasi"]))
            k3.metric("Kas Neto Pendanaan", format_rupiah(totals["pendanaan"]))
            k4.metric("Kenaikan / (Penurunan) Kas", format_rupiah(item["net_change"]))

            statement_view = pd.DataFrame([
                {"Aktivitas": "Operasi", "Arus Kas Neto": totals["operasi"]},
                {"Aktivitas": "Investasi", "Arus Kas Neto": totals["investasi"]},
                {"Aktivitas": "Pendanaan", "Arus Kas Neto": totals["pendanaan"]},
                {"Aktivitas": "Kenaikan / (Penurunan) Neto Kas", "Arus Kas Neto": item["net_change"]},
                {"Aktivitas": "Saldo Kas Awal", "Arus Kas Neto": item["opening_cash"]},
                {"Aktivitas": "Saldo Kas Akhir", "Arus Kas Neto": item["closing_cash"]},
                {"Aktivitas": "Free Cash Flow (estimasi)", "Arus Kas Neto": item["fcf"]},
            ])
            ratios = {}
            trend_df = pd.DataFrame()
            t1, t2, t3, t4, t5 = st.tabs([
                "📋 Aktivitas Arus Kas", "📐 Rasio", "📈 Tren", "🚨 Risiko", "🧠 AI Insights"
            ])
            with t1:
                st.dataframe(statement_view, use_container_width=True, hide_index=True)
                if item["line_items"]:
                    with st.expander(f"📑 Rincian Pos Terklasifikasi ({len(item['line_items'])} baris)", expanded=True):
                        st.dataframe(pd.DataFrame(item["line_items"]), use_container_width=True, hide_index=True)
                if item.get("unclassified_items"):
                    with st.expander(f"⚠️ Pos Belum Terklasifikasi ({len(item['unclassified_items'])} baris)", expanded=False):
                        st.dataframe(pd.DataFrame(item["unclassified_items"]), use_container_width=True, hide_index=True)
                st.caption("Pengelompokan otomatis menggunakan uraian akun. Tinjau kembali klasifikasi setiap baris sebelum memakai hasil sebagai laporan resmi.")
            with t2:
                st.markdown("Rasio yang dapat dihitung dari laporan arus kas dan data pendukung yang tersedia:")
                revenue_rows = data[data[desc_col].astype(str).str.contains("pendapatan|penjualan|revenue", case=False, na=False)]
                revenue_default = to_number(revenue_rows.iloc[0][chosen]) if not revenue_rows.empty else 0.0
                c1, c2 = st.columns(2)
                revenue = c1.number_input("Pendapatan / Penjualan (Rp)", min_value=0.0, value=float(max(0, revenue_default)), key=f"cf_stmt_revenue_{chosen}")
                current_liabilities = c2.number_input("Liabilitas Lancar (Rp)", min_value=0.0, value=0.0, key=f"cf_stmt_current_liab_{chosen}")
                c3, c4 = st.columns(2)
                total_debt = c3.number_input("Total Utang (Rp)", min_value=0.0, value=0.0, key=f"cf_stmt_total_debt_{chosen}")
                debt_service = c4.number_input("Pembayaran Pokok Utang + Bunga (Rp)", min_value=0.0, value=0.0, key=f"cf_stmt_debt_service_{chosen}")
                ratios = {
                    "Cash Flow Margin (%)": round(totals["operasi"] / revenue * 100, 2) if revenue > 0 else "Data pendapatan tidak tersedia",
                    "Operating Cash Flow Ratio (x)": round(totals["operasi"] / current_liabilities, 2) if current_liabilities > 0 else "Data liabilitas lancar tidak tersedia",
                    "Cash Flow to Debt Ratio (%)": round(totals["operasi"] / total_debt * 100, 2) if total_debt > 0 else "Data total utang tidak tersedia",
                    "Cash Flow Coverage (x)": round(totals["operasi"] / debt_service, 2) if debt_service > 0 else "Data pembayaran utang dan bunga tidak tersedia",
                    "Free Cash Flow (Rp)": item["fcf"],
                }
                for label, val in ratios.items():
                    if isinstance(val, (int, float)):
                        metric_card(label, format_rupiah(val) if "(Rp)" in label else f"{val:.2f}" + (" %" if "%" in label else " x"))
                    else:
                        st.info(f"{label}: {val}")
                st.caption("Masukkan data pendukung dari laporan laba rugi dan laporan posisi keuangan. Rasio hanya dihitung jika pembaginya tersedia dan lebih besar dari nol.")
            with t3:
                trend_rows = []
                for period_name, result in saved["results"].items():
                    trend_rows.append({
                        "Periode": str(period_name),
                        "Operasi": result["totals"]["operasi"],
                        "Investasi": result["totals"]["investasi"],
                        "Pendanaan": result["totals"]["pendanaan"],
                        "Perubahan Kas": result["net_change"],
                        "Saldo Kas Akhir": result["closing_cash"]
                    })
                trend_df = pd.DataFrame(trend_rows)
                st.plotly_chart(
                    px.bar(trend_df, x="Periode", y=["Operasi", "Investasi", "Pendanaan"],
                           barmode="group", title="Perbandingan Arus Kas per Aktivitas",
                           color_discrete_sequence=["#FF5C8B", "#7A1F3D", "#F8C8D4"]),
                    use_container_width=True
                )
                st.dataframe(trend_df, use_container_width=True, hide_index=True)
            with t4:
                issues = []
                if totals["operasi"] < 0:
                    issues.append("Arus kas operasi negatif.")
                if item["closing_cash"] is not None and item["closing_cash"] < 0:
                    issues.append("Saldo kas akhir negatif.")
                if item["calculated_closing"] is not None and item["closing_cash"] is not None:
                    difference = item["closing_cash"] - item["calculated_closing"]
                    if abs(difference) > 1:
                        issues.append(f"Rekonsiliasi saldo kas berbeda sebesar {format_rupiah(difference)}.")
                if totals["pendanaan"] > 0 and totals["operasi"] < 0:
                    issues.append("Kas positif berpotensi bergantung pada pendanaan eksternal ketika operasi belum menghasilkan kas.")
                if issues:
                    for issue in issues:
                        st.warning(issue)
                else:
                    st.success("Tidak ditemukan indikator risiko dasar dari pos ringkasan yang tersedia. Ini bukan jaminan tidak ada salah saji.")
            with t5:
                if st.button("🤖 Buat AI Financial Insights", key="ai_cf_statement"):
                    context = (
                        f"Periode: {chosen}\nArus kas operasi: {totals['operasi']:,.0f}\n"
                        f"Arus kas investasi: {totals['investasi']:,.0f}\n"
                        f"Arus kas pendanaan: {totals['pendanaan']:,.0f}\n"
                        f"Perubahan neto kas: {item['net_change']:,.0f}\n"
                        f"Saldo awal: {item['opening_cash']}\nSaldo akhir: {item['closing_cash']}\n"
                        f"Free cash flow estimasi: {item['fcf']:,.0f}"
                    )
                    prompt = (
                        "Analisis arus kas secara profesional. Jelaskan sumber dan penggunaan kas, "
                        "kualitas arus kas operasi, aktivitas investasi dan pendanaan, risiko likuiditas, "
                        "rekonsiliasi saldo, serta rekomendasi yang spesifik. Bedakan fakta angka dan interpretasi. "
                        "Jangan menyimpulkan rasio yang datanya tidak tersedia."
                    )
                    with st.chat_message("assistant"):
                        st.write_stream(ask_gemini_stream(prompt, context, system_role="Analis Arus Kas dan Akuntansi Profesional"))
            summary = {
                "Periode": chosen,
                "Kas Neto Operasi": totals["operasi"],
                "Kas Neto Investasi": totals["investasi"],
                "Kas Neto Pendanaan": totals["pendanaan"],
                "Perubahan Neto Kas": item["net_change"],
                "Saldo Kas Awal": item["opening_cash"],
                "Saldo Kas Akhir": item["closing_cash"],
                "Free Cash Flow": item["fcf"]
            }
            excel_data = make_cf_excel(statement_view, {}, summary, trend_df if len(saved["results"]) > 1 else None)
            st.download_button(
                "⬇️ Download Analisis Cash Flow (Excel)",
                data=excel_data,
                file_name="FINORA_Analisis_Cash_Flow.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            pdf_summary = {**summary, **ratios}
            st.download_button(
                "⬇️ Download Laporan Cash Flow (PDF)",
                data=make_cf_pdf("Laporan Analisis Arus Kas", pdf_summary, statement_view),
                file_name="FINORA_Laporan_Arus_Kas.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        return

    # ------------------- MODE DATA TRANSAKSI / MUTASI BANK -------------------
    st.markdown("### 1. Pemetaan kolom transaksi")
    cols = list(data.columns)
    if not cols:
        st.error("Header atau data transaksi tidak ditemukan.")
        return

    def find_col(keys):
        return next((c for c in cols if any(k in str(c).lower() for k in keys)), None)

    date_auto = find_col(["tanggal", "tgl", "date", "waktu", "periode"]) or cols[0]
    desc_auto = find_col(["keterangan", "deskripsi", "uraian", "nama transaksi", "transaksi", "ket"])
    inflow_auto = find_col(["kas masuk", "penerimaan", "kredit", "credit", "inflow", "debit masuk"])
    outflow_auto = find_col(["kas keluar", "pengeluaran", "debit", "debit keluar", "outflow"])
    amount_auto = find_col(["nominal", "jumlah", "nilai", "amount", "mutasi", "saldo transaksi"])
    category_auto = find_col(["aktivitas", "jenis arus", "kategori", "kelompok"])

    with st.expander("⚙️ Koreksi pemetaan kolom", expanded=True):
        c1, c2, c3 = st.columns(3)
        date_col = c1.selectbox("Tanggal", cols, index=cols.index(date_auto), key="cf_tx_date")
        desc_col = c2.selectbox("Keterangan", ["(Tidak ada)"] + cols, index=(cols.index(desc_auto) + 1 if desc_auto else 0), key="cf_tx_desc")
        category_col = c3.selectbox("Kategori / Aktivitas (opsional)", ["(Deteksi otomatis)"] + cols, index=(cols.index(category_auto) + 1 if category_auto else 0), key="cf_tx_category")
        c4, c5, c6 = st.columns(3)
        inflow_col = c4.selectbox("Kolom kas masuk (opsional)", ["(Tidak ada)"] + cols, index=(cols.index(inflow_auto) + 1 if inflow_auto else 0), key="cf_tx_in")
        outflow_col = c5.selectbox("Kolom kas keluar (opsional)", ["(Tidak ada)"] + cols, index=(cols.index(outflow_auto) + 1 if outflow_auto else 0), key="cf_tx_out")
        amount_col = c6.selectbox("Kolom nominal bertanda +/- (opsional)", ["(Tidak ada)"] + cols, index=(cols.index(amount_auto) + 1 if amount_auto else 0), key="cf_tx_amount")
    opening_cash = st.number_input("Saldo kas awal periode (Rp)", min_value=0.0, value=0.0, step=1000000.0, key="cf_opening_cash")
    st.caption("Isi kolom kas masuk dan kas keluar jika tersedia. Jika hanya ada satu kolom nominal, gunakan angka positif untuk masuk dan negatif untuk keluar.")

    if st.button("📊 JALANKAN ANALISIS ARUS KAS", use_container_width=True, key="run_cf_transactions"):
        work = data.copy()
        work["_tanggal"] = pd.to_datetime(work[date_col], errors="coerce", dayfirst=True)
        if work["_tanggal"].isna().all() and pd.api.types.is_numeric_dtype(work[date_col]):
            work["_tanggal"] = pd.to_datetime(work[date_col], errors="coerce", unit="D", origin="1899-12-30")
        if work["_tanggal"].isna().all():
            st.error("Kolom tanggal tidak dapat dibaca. Periksa kembali pemetaan tanggal.")
            return

        if inflow_col != "(Tidak ada)" and outflow_col != "(Tidak ada)":
            work["_kas_masuk"] = work[inflow_col].apply(to_number).clip(lower=0)
            work["_kas_keluar"] = work[outflow_col].apply(to_number).abs()
        elif amount_col != "(Tidak ada)":
            signed = work[amount_col].apply(to_number)
            work["_kas_masuk"] = signed.clip(lower=0)
            work["_kas_keluar"] = signed.clip(upper=0).abs()
        else:
            st.error("Pilih pasangan kolom kas masuk/kas keluar atau satu kolom nominal bertanda positif/negatif.")
            return

        work = work.dropna(subset=["_tanggal"]).sort_values("_tanggal").copy()
        work["_net"] = work["_kas_masuk"] - work["_kas_keluar"]
        if desc_col != "(Tidak ada)":
            work["_keterangan"] = work[desc_col].astype(str)
        else:
            work["_keterangan"] = "Transaksi"

        def guess_activity(row):
            if category_col != "(Deteksi otomatis)":
                raw_cat = str(row[category_col]).lower()
            else:
                raw_cat = str(row["_keterangan"]).lower()
            if any(k in raw_cat for k in ["invest", "aset tetap", "mesin", "kendaraan", "properti", "deposito", "saham"]):
                return "Investasi"
            if any(k in raw_cat for k in ["pinjaman", "utang bank", "modal", "saham baru", "dividen", "pelunasan pokok"]):
                return "Pendanaan"
            return "Operasi"

        work["_aktivitas"] = work.apply(guess_activity, axis=1)
        work["_saldo_berjalan"] = opening_cash + work["_net"].cumsum()
        work["_bulan"] = work["_tanggal"].dt.to_period("M").astype(str)

        totals = {}
        for act in ["Operasi", "Investasi", "Pendanaan"]:
            part = work[work["_aktivitas"] == act]
            totals[act] = float(part["_net"].sum())
        total_in = float(work["_kas_masuk"].sum())
        total_out = float(work["_kas_keluar"].sum())
        net_change = total_in - total_out
        closing = opening_cash + net_change
        month_count = max(1, work["_bulan"].nunique())
        avg_operating_outflow = float(work[(work["_aktivitas"] == "Operasi")]["_kas_keluar"].sum()) / month_count
        fcf = totals["Operasi"] - float(work[(work["_aktivitas"] == "Investasi")]["_kas_keluar"].sum())

        monthly = work.groupby("_bulan", as_index=False).agg(
            Kas_Masuk=("_kas_masuk", "sum"),
            Kas_Keluar=("_kas_keluar", "sum"),
            Arus_Kas_Neto=("_net", "sum")
        )
        monthly["Saldo_Akhir_Estimasi"] = opening_cash + monthly["Arus_Kas_Neto"].cumsum()
        outflows = work[work["_kas_keluar"] > 0]
        anomalies = pd.DataFrame()
        if len(outflows) >= 4:
            q1 = outflows["_kas_keluar"].quantile(.25)
            q3 = outflows["_kas_keluar"].quantile(.75)
            upper = q3 + 1.5 * (q3 - q1)
            anomalies = outflows[outflows["_kas_keluar"] > upper].copy()

        result = {
            "df": work, "monthly": monthly, "totals": totals,
            "total_in": total_in, "total_out": total_out, "net_change": net_change,
            "opening": opening_cash, "closing": closing, "fcf": fcf,
            "avg_operating_outflow": avg_operating_outflow, "months": month_count,
            "anomalies": anomalies, "filename": uploaded_cf.name
        }
        st.session_state["cf_transaction_results"] = result
        st.success(f"Analisis transaksi berhasil dijalankan ({len(work)} transaksi terbaca).")

    if "cf_transaction_results" in st.session_state:
        res = st.session_state["cf_transaction_results"]
        st.markdown("---")
        st.markdown("### 2. Cash Flow Overview")
        a, b, c, d = st.columns(4)
        a.metric("Total Kas Masuk", format_rupiah(res["total_in"]))
        b.metric("Total Kas Keluar", format_rupiah(res["total_out"]))
        c.metric("Arus Kas Neto", format_rupiah(res["net_change"]))
        d.metric("Saldo Kas Akhir", format_rupiah(res["closing"]))

        tab_summary, tab_activity, tab_ratios, tab_trends, tab_risk, tab_ai = st.tabs([
            "📊 Ringkasan", "🏢 Aktivitas", "📐 Rasio", "📈 Tren", "🚨 Risiko", "🧠 AI Insights"
        ])
        with tab_summary:
            st.dataframe(pd.DataFrame([
                {"Komponen": "Saldo kas awal", "Nilai (Rp)": res["opening"]},
                {"Komponen": "Total kas masuk", "Nilai (Rp)": res["total_in"]},
                {"Komponen": "Total kas keluar", "Nilai (Rp)": -res["total_out"]},
                {"Komponen": "Kenaikan / (Penurunan) neto kas", "Nilai (Rp)": res["net_change"]},
                {"Komponen": "Saldo kas akhir", "Nilai (Rp)": res["closing"]},
            ]), use_container_width=True, hide_index=True)
            st.caption("Saldo kas akhir dihitung dari saldo awal ditambah kas masuk dikurangi kas keluar.")
        with tab_activity:
            activity_df = pd.DataFrame([
                {"Aktivitas": k, "Arus Kas Neto (Rp)": v} for k, v in res["totals"].items()
            ])
            st.dataframe(activity_df, use_container_width=True, hide_index=True)
            st.plotly_chart(px.bar(activity_df, x="Aktivitas", y="Arus Kas Neto (Rp)",
                                   title="Arus Kas Neto per Aktivitas",
                                   color_discrete_sequence=["#FF5C8B"]), use_container_width=True)
            st.caption("Klasifikasi otomatis berdasarkan kata kunci pada keterangan/kategori. Periksa dan koreksi kategori bila diperlukan.")
        with tab_ratios:
            rc1, rc2 = st.columns(2)
            revenue = rc1.number_input("Pendapatan / Penjualan (Rp)", min_value=0.0, value=0.0, key="cf_tx_revenue")
            current_liabilities = rc2.number_input("Liabilitas Lancar (Rp)", min_value=0.0, value=0.0, key="cf_tx_current_liab")
            rc3, rc4 = st.columns(2)
            total_debt = rc3.number_input("Total Utang (Rp)", min_value=0.0, value=0.0, key="cf_tx_total_debt")
            debt_service = rc4.number_input("Pembayaran Pokok Utang + Bunga (Rp)", min_value=0.0, value=0.0, key="cf_tx_debt_service")
            ratios = {
                "Cash Flow Margin (%)": round(res["totals"]["Operasi"] / revenue * 100, 2) if revenue > 0 else "Data pendapatan tidak tersedia",
                "Operating Cash Flow Ratio (x)": round(res["totals"]["Operasi"] / current_liabilities, 2) if current_liabilities > 0 else "Data liabilitas lancar tidak tersedia",
                "Cash Flow to Debt Ratio (%)": round(res["totals"]["Operasi"] / total_debt * 100, 2) if total_debt > 0 else "Data total utang tidak tersedia",
                "Cash Flow Coverage (x)": round(res["totals"]["Operasi"] / debt_service, 2) if debt_service > 0 else "Data pembayaran utang dan bunga tidak tersedia",
                "Free Cash Flow (Rp)": res["fcf"]
            }
            for ratio_name, ratio_value in ratios.items():
                if isinstance(ratio_value, (int, float)):
                    metric_card(ratio_name, format_rupiah(ratio_value) if "(Rp)" in ratio_name else f"{ratio_value:.2f}" + (" %" if "%" in ratio_name else " x"))
                else:
                    st.info(f"{ratio_name}: {ratio_value}")
            st.metric("Rata-rata Pengeluaran Operasi per Bulan", format_rupiah(res["avg_operating_outflow"]))
            st.markdown("**Rumus yang digunakan jika data tersedia:**")
            st.markdown(
                "- Cash Flow Margin = Arus Kas Operasi ÷ Penjualan × 100%\n"
                "- Operating Cash Flow Ratio = Arus Kas Operasi ÷ Liabilitas Lancar\n"
                "- Cash Flow to Debt = Arus Kas Operasi ÷ Total Utang\n"
                "- Free Cash Flow = Arus Kas Operasi − Belanja Modal (CapEx)"
            )
        with tab_trends:
            st.plotly_chart(px.line(res["monthly"], x="_bulan", y=["Kas_Masuk", "Kas_Keluar"],
                                    markers=True, title="Tren Kas Masuk dan Kas Keluar",
                                    color_discrete_sequence=["#FF5C8B", "#7A1F3D"]),
                            use_container_width=True)
            st.plotly_chart(px.area(res["monthly"], x="_bulan", y="Saldo_Akhir_Estimasi",
                                    title="Perkembangan Estimasi Saldo Kas",
                                    color_discrete_sequence=["#FF5C8B"]), use_container_width=True)
            st.dataframe(res["monthly"], use_container_width=True, hide_index=True)
        with tab_risk:
            risk_messages = []
            if res["totals"]["Operasi"] < 0:
                risk_messages.append("Arus kas operasi negatif: kegiatan utama belum menghasilkan kas bersih.")
            if res["closing"] < 0:
                risk_messages.append("Saldo kas akhir negatif: periksa saldo awal dan kelengkapan pencatatan.")
            if res["totals"]["Pendanaan"] > 0 and res["totals"]["Operasi"] < 0:
                risk_messages.append("Kas masuk dari pendanaan menutup kondisi arus kas operasi yang negatif.")
            if not res["anomalies"].empty:
                risk_messages.append(f"Terdapat {len(res['anomalies'])} transaksi pengeluaran ekstrem berdasarkan metode IQR.")
                st.dataframe(res["anomalies"][["_tanggal", "_keterangan", "_kas_keluar", "_aktivitas"]], use_container_width=True)
            if risk_messages:
                for msg in risk_messages:
                    st.warning(msg)
            else:
                st.success("Tidak ditemukan indikator risiko dasar berdasarkan pengujian otomatis.")
            st.caption("Deteksi anomali adalah indikator untuk pemeriksaan lanjutan, bukan bukti kecurangan.")
        with tab_ai:
            if st.button("🤖 Buat AI Financial Insights", key="ai_cf_transaction"):
                context = (
                    f"Kas masuk: {res['total_in']:,.0f}\nKas keluar: {res['total_out']:,.0f}\n"
                    f"Saldo awal: {res['opening']:,.0f}\nSaldo akhir: {res['closing']:,.0f}\n"
                    f"Arus kas operasi: {res['totals']['Operasi']:,.0f}\n"
                    f"Arus kas investasi: {res['totals']['Investasi']:,.0f}\n"
                    f"Arus kas pendanaan: {res['totals']['Pendanaan']:,.0f}\n"
                    f"Free cash flow estimasi: {res['fcf']:,.0f}\n"
                    f"Rata-rata pengeluaran operasi bulanan: {res['avg_operating_outflow']:,.0f}\n"
                    f"Jumlah transaksi anomali: {len(res['anomalies'])}"
                )
                prompt = (
                    "Buat analisis arus kas yang mencakup sumber dan penggunaan kas, "
                    "kemampuan operasi menghasilkan kas, investasi, pendanaan, tren, risiko likuiditas, "
                    "dan rekomendasi praktis. Sebutkan keterbatasan karena klasifikasi berasal dari kata kunci. "
                    "Jangan membuat angka atau rasio yang tidak tersedia."
                )
                with st.chat_message("assistant"):
                    st.write_stream(ask_gemini_stream(prompt, context, system_role="FINORA AI Cash Flow Analyst"))

        export_df = res["df"][["_tanggal", "_keterangan", "_aktivitas", "_kas_masuk", "_kas_keluar", "_net", "_saldo_berjalan"]].copy()
        summary = {
            "Saldo kas awal": res["opening"], "Kas masuk": res["total_in"],
            "Kas keluar": res["total_out"], "Arus kas neto": res["net_change"],
            "Arus kas operasi": res["totals"]["Operasi"],
            "Arus kas investasi": res["totals"]["Investasi"],
            "Arus kas pendanaan": res["totals"]["Pendanaan"],
            "Free Cash Flow": res["fcf"], "Saldo kas akhir": res["closing"]
        }
        excel_file = make_cf_excel(export_df, ratios, summary, res["monthly"], res["anomalies"])
        st.download_button(
            "⬇️ Download Analisis Arus Kas (Excel)",
            data=excel_file,
            file_name="FINORA_Analisis_Arus_Kas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        pdf_summary = {**summary, **ratios}
        st.download_button(
            "⬇️ Download Laporan Arus Kas (PDF)",
            data=make_cf_pdf("Laporan Analisis Arus Kas", pdf_summary, export_df),
            file_name="FINORA_Laporan_Arus_Kas.pdf",
            mime="application/pdf",
            use_container_width=True
        )

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
