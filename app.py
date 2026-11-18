import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import json
import os
import io
import base64
import hashlib
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from google import genai

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
.metric-card {
    background: white;
    padding: 16px;
    border-radius: 12px;
    border-left: 4px solid #FF5C8B;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    margin-bottom: 15px;
}
.metric-card p { color: #555; font-size: 14px; margin: 0; font-weight: 600; }
.metric-card h3 { color: #7A1F3D; margin: 5px 0 0 0; font-size: 24px; font-weight: 700; }
.insight-box {
    background-color: white;
    padding: 20px;
    border-radius: 10px;
    border-left: 5px solid #FF5C8B;
    margin-bottom: 15px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.03);
}
.insight-box h4 { margin-top: 0; color: #7A1F3D; font-size: 18px; margin-bottom: 10px; }
.insight-box p { margin-bottom: 8px; font-size: 14px; line-height: 1.6; color: #333; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# FUNGSI PEMBANTU FORMAT ANGKA & KARTU
# ============================================================
def format_rupiah(val):
    if abs(val) >= 1_000_000_000_000: return f"Rp {val/1_000_000_000_000:,.2f} T"
    elif abs(val) >= 1_000_000_000: return f"Rp {val/1_000_000_000:,.2f} M"
    elif abs(val) >= 1_000_000: return f"Rp {val/1_000_000:,.2f} Jt"
    else: return f"Rp {val:,.0f}"

def metric_card(label, value, help_text=""):
    st.markdown(f"""
    <div class="metric-card" title="{help_text}">
        <p>{label}</p>
        <h3>{value}</h3>
    </div>
    """, unsafe_allow_html=True)

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
    c.execute("INSERT INTO files(filename, file_hash, uploaded_at) VALUES (?,?,?)",
              (filename, file_hash, datetime.now().isoformat()))
    fid = c.lastrowid; conn.commit(); conn.close()
    return fid

def save_analysis(file_id, mode, selected_modules, summary):
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO analyses(file_id, mode, selected_modules, status, summary_json, created_at) VALUES (?,?,?,?,?,?)",
              (file_id, mode, json.dumps(selected_modules), "completed", json.dumps(summary, default=str), datetime.now().isoformat()))
    aid = c.lastrowid; conn.commit(); conn.close()
    return aid

def save_findings(analysis_id, findings):
    if not findings:
        return
    conn = get_conn(); c = conn.cursor()
    for f in findings:
        c.execute("INSERT INTO findings(analysis_id, module, risk_level, title, detail, recommendation) VALUES (?,?,?,?,?,?)",
                  (analysis_id, f["module"], f["risk_level"], f["title"], f["detail"], f["recommendation"]))
    conn.commit(); conn.close()

def fetch_analyses():
    conn = get_conn()
    df = pd.read_sql_query("SELECT a.id, f.filename, a.mode, a.created_at FROM analyses a JOIN files f ON a.file_id = f.id ORDER BY a.created_at DESC", conn)
    conn.close(); return df

# ============================================================
# AI ENGINE
# ============================================================
def get_gemini_client():
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    except Exception:
        api_key = os.environ.get("GEMINI_API_KEY", "")
        
    if not api_key:
        return None
    try:
        return genai.Client(api_key=api_key)
    except Exception:
        return None

def ask_gemini_stream(prompt, audit_context, system_role="Asisten Auditor Profesional"):
    try:
        client = get_gemini_client()
        if client is None:
            yield "⚠️ API Key Gemini belum diatur di Secrets Streamlit."
            return

        model_id = "gemini-1.5-flash"

        full_prompt = f"""
        Kamu adalah FINORA AI, {system_role}.
        Jawab pertanyaan berdasarkan konteks data yang diberikan secara ringkas, padat, dan profesional.

        [KONTEKS DATA]
        {audit_context}

        [PERTANYAAN / INSTRUKSI]
        {prompt}
        """

        response_stream = client.models.generate_content_stream(model=model_id, contents=full_prompt)
        for chunk in response_stream:
            if getattr(chunk, "text", None): 
                yield chunk.text

    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan: {str(e)}"

# ============================================================
# AUDIT ENGINE V5
# ============================================================
def auto_detect_columns(df):
    mapping = {"date": None, "amount": None, "desc": None, "debit": None, "credit": None, "budget": None, "invoice": None, "cash_in": None, "cash_out": None}
    for col in df.columns:
        c = str(col).lower()
        if not mapping["date"] and any(k in c for k in ["tanggal", "tgl", "date", "waktu", "periode"]): mapping["date"] = col
        if not mapping["amount"] and any(k in c for k in ["actual", "realisasi", "nominal", "jumlah", "total", "nilai", "amount", "harga"]): mapping["amount"] = col
        if not mapping["budget"] and any(k in c for k in ["anggaran", "budget", "pagu"]): mapping["budget"] = col
        if not mapping["desc"] and any(k in c for k in ["deskripsi", "keterangan", "uraian", "category", "kategori", "nama", "item", "barang"]): mapping["desc"] = col
        if not mapping["debit"] and any(k in c for k in ["debit", "masuk", "pemasukan", "in"]): mapping["debit"] = col
        if not mapping["credit"] and any(k in c for k in ["kredit", "keluar", "pengeluaran", "out"]): mapping["credit"] = col
        if not mapping["cash_in"] and any(k in c for k in ["kas masuk", "pemasukan", "cash in", "debit"]): mapping["cash_in"] = col
        if not mapping["cash_out"] and any(k in c for k in ["kas keluar", "pengeluaran", "cash out", "kredit"]): mapping["cash_out"] = col
        if not mapping["invoice"] and any(k in c for k in ["invoice", "faktur", "nota", "ref", "no."]): mapping["invoice"] = col
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
            dates = pd.to_datetime(df[date_col], errors="coerce")
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
# CASH FLOW INTELLIGENCE ENGINE (FITUR BARU SESUAI KONSEP)
# ============================================================
def process_cash_flow_data(df, date_col, desc_col, in_col, out_col, initial_balance=0.0):
    pdf = df.copy()
    
    # 1. Konversi Tanggal & Urutkan
    if date_col:
        pdf[date_col] = pd.to_datetime(pdf[date_col], errors="coerce")
        pdf = pdf.sort_values(by=date_col).reset_index(drop=True)
        pdf['Bulan'] = pdf[date_col].dt.to_period('M').astype(str)
    else:
        pdf['Bulan'] = 'Tidak Diketahui'

    # 2. Ambil Nilai Masuk & Keluar
    pdf['Kas Masuk'] = pd.to_numeric(pdf[in_col], errors="coerce").fillna(0.0) if in_col else 0.0
    pdf['Kas Keluar'] = pd.to_numeric(pdf[out_col], errors="coerce").fillna(0.0) if out_col else 0.0
    pdf['Net Cash Flow'] = pdf['Kas Masuk'] - pdf['Kas Keluar']

    # 3. Running Balance (Saldo Berjalan)
    pdf['Running Balance'] = initial_balance + pdf['Net Cash Flow'].cumsum()

    # 4. Kategorisasi Otomatis Berdasarkan Kata Kunci Deskripsi
    def auto_categorize(text):
        t = str(text).lower()
        if any(k in t for k in ["bensin", "bbm", "tol", "parkir", "transport", "travel", "akomodasi"]): return "Transportasi & Perjalanan"
        elif any(k in t for k in ["gaji", "honor", "tunjangan", "payroll", "upah", "salary"]): return "Gaji & SDM"
        elif any(k in t for k in ["listrik", "air", "internet", "wifi", "telkom", "pulsa", "utility", "pln"]): return "Utilitas & Kantor"
        elif any(k in t for k in ["sewa", "rent", "gedung", "kantor"]): return "Sewa & Fasilitas"
        elif any(k in t for k in ["makan", "konsumsi", "snack", "f&b", "restoran"]): return "Konsumsi"
        elif any(k in t for k in ["pajak", "tax", "retribusi"]): return "Pajak & Legal"
        elif any(k in t for k in ["penjualan", "sales", "pendapatan", "revenue", "terima", "inflow", "setoran"]): return "Penerimaan / Penjualan"
        else: return "Operasional Lainnya"

    if desc_col:
        pdf['Kategori Pengeluaran'] = pdf[desc_col].apply(auto_categorize)
    else:
        pdf['Kategori Pengeluaran'] = "Umum"

    # 5. Hitung Burn Rate & Runway
    total_out = pdf['Kas Keluar'].sum()
    if date_col and not pdf[date_col].dropna().empty:
        min_date = pdf[date_col].min()
        max_date = pdf[date_col].max()
        diff_days = (max_date - min_date).days
        months_count = max(1, diff_days / 30.44)
    else:
        months_count = 1.0

    monthly_burn_rate = total_out / months_count
    current_saldo = pdf['Running Balance'].iloc[-1] if not pdf.empty else initial_balance
    cash_runway_months = round(current_saldo / monthly_burn_rate, 1) if monthly_burn_rate > 0 else 99.9

    # 6. Deteksi Anomali Kas (Pengeluaran > 3x Rata-rata)
    out_only = pdf[pdf['Kas Keluar'] > 0]['Kas Keluar']
    anomalies_df = pd.DataFrame()
    if not out_only.empty:
        mean_out = out_only.mean()
        std_out = out_only.std() if len(out_only) > 1 else 0
        threshold = mean_out + (2 * std_out) if std_out > 0 else mean_out * 3
        anomalies_df = pdf[pdf['Kas Keluar'] > threshold]

    metrics = {
        "total_masuk": pdf['Kas Masuk'].sum(),
        "total_keluar": total_out,
        "net_cash_flow": pdf['Net Cash Flow'].sum(),
        "saldo_akhir": current_saldo,
        "monthly_burn_rate": monthly_burn_rate,
        "cash_runway_months": cash_runway_months
    }

    return pdf, metrics, anomalies_df

# ============================================================
# SMART FINANCIAL ENGINE & EXPORT
# ============================================================
def auto_extract_financials(df, entity_type):
    results = {
        "pendapatan": 0.0, "laba": 0.0, "aset_lancar": 0.0, "liabilitas_lancar": 0.0,
        "total_aset": 0.0, "total_liabilitas": 0.0, "ekuitas": 0.0,
        "real_pendapatan": 0.0, "ang_pendapatan": 0.0,
        "real_belanja": 0.0, "ang_belanja": 0.0,
        "surplus": 0.0, "pad": 0.0
    }
    
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
            if mask_exact.any():
                return pd.to_numeric(df.loc[mask_exact, target_col], errors='coerce').fillna(0).sum()
        if contains_keys:
            mask_contains = clean_names.str.contains('|'.join(contains_keys), na=False)
            if mask_contains.any():
                return pd.to_numeric(df.loc[mask_contains, target_col], errors='coerce').fillna(0).sum()
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

        results.update({
            "real_pendapatan": real_p, "ang_pendapatan": ang_p,
            "real_belanja": real_b, "ang_belanja": ang_b,
            "pad": pad, "surplus": surplus
        })
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
        real_p = nilai.get("real_pendapatan", 0)
        ang_p = nilai.get("ang_pendapatan", 0)
        real_b = nilai.get("real_belanja", 0)
        ang_b = nilai.get("ang_belanja", 0)
        pad = nilai.get("pad", 0)
        
        rasio["Efektivitas Pendapatan (%)"] = round((real_p / ang_p) * 100, 2) if ang_p > 0 else 0.0
        rasio["Efisiensi Belanja (%)"] = round((real_b / ang_b) * 100, 2) if ang_b > 0 else 0.0
        rasio["Kemandirian Daerah (%)"] = round((pad / real_p) * 100, 2) if real_p > 0 else 0.0
    else:
        pendapatan = nilai.get("pendapatan", 0)
        laba = nilai.get("laba", 0)
        aset_lancar = nilai.get("aset_lancar", 0)
        liab_lancar = nilai.get("liabilitas_lancar", 0)
        total_aset = nilai.get("total_aset", 0)
        total_liab = nilai.get("total_liabilitas", 0)
        ekuitas = nilai.get("ekuitas", 0)

        rasio["Current Ratio (x)"] = round(aset_lancar / liab_lancar, 2) if liab_lancar > 0 else 0.0
        rasio["Net Profit Margin (%)"] = round((laba / pendapatan) * 100, 2) if pendapatan > 0 else 0.0
        rasio["ROA (%)"] = round((laba / total_aset) * 100, 2) if total_aset > 0 else 0.0
        rasio["ROE (%)"] = round((laba / ekuitas) * 100, 2) if ekuitas > 0 else 0.0
        rasio["Debt to Equity (x)"] = round(total_liab / ekuitas, 2) if ekuitas > 0 else 0.0
    return rasio

def get_detailed_insight(ratio_name, val):
    return f"<b>Kondisi:</b> Nilai indikator {ratio_name} tercatat sebesar {val}.<br><b>Rekomendasi:</b> Evaluasi berkala sesuai standar operasional keuangan."

def generate_pdf_report(filename, entity_type, summary_data, findings_list, rasio_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#7A1F3D"), spaceAfter=6)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=10, textColor=colors.grey, spaceAfter=16)
    h2_style = ParagraphStyle('H2Style', parent=styles['Heading2'], fontSize=13, textColor=colors.HexColor("#7A1F3D"), spaceBefore=12, spaceAfter=8)
    normal_style = styles['Normal']

    elements = []
    elements.append(Paragraph("LAPORAN HASIL AUDIT & KEUANGAN FINORA AI", title_style))
    elements.append(Paragraph(f"File: {filename} | Entitas: {entity_type} | Tanggal: {datetime.now().strftime('%d %B %Y')}", sub_style))

    elements.append(Paragraph("1. Ringkasan Eksekutif", h2_style))
    summary_table_data = [["Metrik", "Nilai"]]
    for k, v in summary_data.items():
        val_str = format_rupiah(v) if isinstance(v, (int, float)) else str(v)
        summary_table_data.append([str(k), val_str])

    t_sum = Table(summary_table_data, colWidths=[250, 240])
    t_sum.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#7A1F3D")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(t_sum)
    elements.append(Spacer(1, 12))

    if rasio_dict:
        elements.append(Paragraph("2. Analisis Rasio Keuangan", h2_style))
        rasio_table_data = [["Indikator Rasio", "Hasil"]]
        for k, v in rasio_dict.items():
            rasio_table_data.append([k, f"{v}"])
        t_rasio = Table(rasio_table_data, colWidths=[250, 240])
        t_rasio.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#FF5C8B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(t_rasio)
        elements.append(Spacer(1, 12))

    elements.append(Paragraph("3. Rincian Temuan Audit", h2_style))
    if not findings_list:
        elements.append(Paragraph("Tidak ditemukan anomali signifikan.", normal_style))
    else:
        find_table_data = [["Modul", "Risiko", "Temuan", "Rekomendasi"]]
        for f in findings_list:
            find_table_data.append([
                Paragraph(f["module"], normal_style),
                Paragraph(f["risk_level"], normal_style),
                Paragraph(f["title"], normal_style),
                Paragraph(f["recommendation"], normal_style)
            ])
        t_find = Table(find_table_data, colWidths=[100, 65, 160, 165])
        t_find.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#7A1F3D")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('PADDING', (0, 0), (-1, -1), 'TOP'),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(t_find)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# MAIN UI APPLICATION
# ============================================================
logo_html = f'<img src="data:image/png;base64,{LOGO_B64}" width="42" style="margin-right:12px; border-radius:8px;">' if LOGO_B64 else "📊 "
st.markdown(f"""
<div class="finora-header">
    <h1>{logo_html}FINORA AI — Pro Audit & Financial Intelligence</h1>
    <p>Sistem Audit Otomatis, Cash Flow Intelligence, Analisis Rasio & Asisten Auditor AI</p>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Pengaturan & Modul")
    entity_type = st.selectbox(
        "Jenis Entitas",
        ["Perusahaan / UMKM (Komersial)", "Instansi Pemerintahan / Sektor Publik"]
    )
    audit_mode = st.radio(
        "Mode Analisis Utama",
        ["Audit Transaksi & Jurnal", "Analisis Laporan Keuangan", "Cash Flow Intelligence (Arus Kas)", "Kombinasi Lengkap"]
    )
    
    st.markdown("---")
    st.subheader("🧩 Modul Audit Aktif")
    mod_empty = st.checkbox("Kelengkapan Data (Kosong/Null)", value=True)
    mod_dup = st.checkbox("Deteksi Duplikasi (Double Entry)", value=True)
    mod_neg = st.checkbox("Observasi Nilai Negatif", value=True)
    mod_over = st.checkbox("Cek Overbudget (Realisasi > Pagu)", value=True)
    mod_wknd = st.checkbox("Transaksi Akhir Pekan (Weekend)", value=True)
    mod_word = st.checkbox("Kata Kunci Sensitif (Void/Adj)", value=True)
    mod_iqr = st.checkbox("Deteksi Outlier Statistik (IQR)", value=True)

    st.markdown("---")
    st.subheader("🗂️ Riwayat Analisis")
    history_df = fetch_analyses()
    if not history_df.empty:
        selected_hist = st.selectbox(
            "Lihat Riwayat Tersimpan",
            options=["-- Pilih Riwayat --"] + [f"#{r['id']} - {r['filename']} ({r['created_at'][:10]})" for _, r in history_df.iterrows()]
        )
    else:
        st.caption("Belum ada riwayat tersimpan.")

# --- FILE UPLOADER ---
uploaded_file = st.file_uploader("📂 Unggah File Laporan Keuangan / Mutasi Kas / Buku Besar (Excel / CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_id = save_file(uploaded_file.name, file_bytes)

    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
            sheet_name = st.selectbox("Pilih Sheet Excel:", excel_file.sheet_names) if len(excel_file.sheet_names) > 1 else excel_file.sheet_names[0]
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    st.success(f"Berhasil memuat **{uploaded_file.name}** ({len(df):,} baris, {len(df.columns)} kolom).")

    with st.expander("👁️ Pratinjau Data Mentah (10 Baris Pertama)", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)

    # Deteksi & Mapping Kolom
    auto_map = auto_detect_columns(df)
    cols_options = [None] + list(df.columns)

    with st.expander("🛠️ Pemetaan Kolom Otomatis (Sesuaikan jika perlu)", expanded=False):
        c1, c2, c3, c4 = st.columns(4)
        col_date = c1.selectbox("Kolom Tanggal", cols_options, index=cols_options.index(auto_map["date"]) if auto_map["date"] in cols_options else 0)
        col_amount = c2.selectbox("Kolom Nominal / Realisasi", cols_options, index=cols_options.index(auto_map["amount"]) if auto_map["amount"] in cols_options else 0)
        col_budget = c3.selectbox("Kolom Anggaran / Pagu", cols_options, index=cols_options.index(auto_map["budget"]) if auto_map["budget"] in cols_options else 0)
        col_desc = c4.selectbox("Kolom Deskripsi / Akun", cols_options, index=cols_options.index(auto_map["desc"]) if auto_map["desc"] in cols_options else 0)

        c5, c6 = st.columns(2)
        col_cash_in = c5.selectbox("Kolom Kas Masuk / Debit (Opsional)", cols_options, index=cols_options.index(auto_map["cash_in"]) if auto_map["cash_in"] in cols_options else 0)
        col_cash_out = c6.selectbox("Kolom Kas Keluar / Kredit (Opsional)", cols_options, index=cols_options.index(auto_map["cash_out"]) if auto_map["cash_out"] in cols_options else 0)

    # Jalankan Engine Audit
    all_findings = []
    flagged_tables = {}
    selected_modules = []

    if mod_empty:
        selected_modules.append("Kelengkapan Data")
        f, flg = audit_empty(df, [col_date, col_amount, col_desc])
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Data Kosong"] = flg

    if mod_dup:
        selected_modules.append("Data Duplikat")
        dup_cols = [c for c in [col_date, col_amount, col_desc] if c]
        f, flg = audit_duplicates(df, dup_cols)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Transaksi Duplikat"] = flg

    if mod_neg:
        selected_modules.append("Nilai Negatif")
        f, flg = audit_negative(df, col_amount)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Nilai Negatif"] = flg

    if mod_over:
        selected_modules.append("Overbudget")
        f, flg = audit_overbudget(df, col_budget, col_amount)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Overbudget"] = flg

    if mod_wknd:
        selected_modules.append("Transaksi Weekend")
        f, flg = audit_weekend(df, col_date)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Transaksi Weekend"] = flg

    if mod_word:
        selected_modules.append("Kata Kunci Mencurigakan")
        f, flg = audit_suspicious_words(df, col_desc)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Kata Kunci Mencurigakan"] = flg

    if mod_iqr:
        selected_modules.append("Outlier IQR")
        f, flg = audit_outliers_iqr(df, col_amount)
        all_findings.extend(f)
        if not flg.empty: flagged_tables["Outlier Statistik"] = flg

    # Jalankan Cash Flow Intelligence Engine
    initial_cash = st.sidebar.number_input("Saldo Awal Kas (Rp)", value=0.0, step=1_000_000.0, format="%f")
    cf_df, cf_metrics, cf_anomalies = process_cash_flow_data(df, col_date, col_desc, col_cash_in or col_amount, col_cash_out or col_amount, initial_cash)

    # Ekstraksi Keuangan & Rasio
    fin_values = auto_extract_financials(df, entity_type)
    rasio_hasil = hitung_rasio_dari_nilai(fin_values, entity_type)

    # TABS UTAMA (DITAMBAHKAN TAB CASH FLOW)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🔍 Hasil Temuan Audit",
        "💵 Cash Flow Intelligence",
        "📈 Analisis Rasio & Keuangan",
        "🤖 Konsultasi Auditor AI",
        "📥 Simpan & Ekspor Laporan"
    ])

    # ---------------- TAB 1: TEMUAN AUDIT ----------------
    with tab1:
        high_risk = sum(1 for x in all_findings if x["risk_level"] == "Tinggi")
        med_risk = sum(1 for x in all_findings if x["risk_level"] == "Sedang")
        low_risk = sum(1 for x in all_findings if x["risk_level"] == "Rendah")

        m1, m2, m3, m4 = st.columns(4)
        with m1: metric_card("Total Baris Diperiksa", f"{len(df):,}")
        with m2: metric_card("Temuan Risiko Tinggi", f"{high_risk}")
        with m3: metric_card("Temuan Risiko Sedang", f"{med_risk}")
        with m4: metric_card("Observasi Risiko Rendah", f"{low_risk}")

        if not all_findings:
            st.success("🎉 Tidak ditemukan anomali atau pelanggaran aturan pada modul audit yang dipilih.")
        else:
            st.subheader("Daftar Temuan Audit")
            for idx, item in enumerate(all_findings, 1):
                badge = "🔴" if item["risk_level"] == "Tinggi" else ("🟠" if item["risk_level"] == "Sedang" else "🟡")
                st.markdown(f"""
                <div class="insight-box">
                    <h4>{badge} [{item['risk_level']}] {item['module']} — {item['title']}</h4>
                    <p><b>Detail:</b> {item['detail']}</p>
                    <p><b>Rekomendasi Audit:</b> {item['recommend'] if 'recommend' in item else item['recommendation']}</p>
                </div>
                """, unsafe_allow_html=True)

            st.subheader("📋 Detail Baris Terindikasi (Flagged Rows)")
            for tbl_name, tbl_df in flagged_tables.items():
                with st.expander(f"Lihat Data: {tbl_name} ({len(tbl_df)} baris)"):
                    st.dataframe(tbl_df, use_container_width=True)

    # ---------------- TAB 2: CASH FLOW INTELLIGENCE (BARU) ----------------
    with tab2:
        st.subheader("💵 Analisis Arus Kas & Ketahanan Likuiditas (Cash Flow Intelligence)")
        st.markdown("Fitur ini menganalisis saldo berjalan (*running balance*), tingkat pembakaran kas (*burn rate*), ketahanan kas (*cash runway*), serta deteksi anomali pengeluaran otomatis.")

        # Kartu Metrik Likuiditas
        c1, c2, c3, c4 = st.columns(4)
        with c1: metric_card("Saldo Kas Saat Ini", format_rupiah(cf_metrics["saldo_akhir"]), "Estimasi posisi kas terakhir berdasarkan transaksi masuk & keluar")
        with c2: metric_card("Total Kas Masuk", format_rupiah(cf_metrics["total_masuk"]), "Akumulasi seluruh pemasukan")
        with c3: metric_card("Total Kas Keluar", format_rupiah(cf_metrics["total_keluar"]), "Akumulasi seluruh pengeluaran")
        with c4: 
            runway_str = f"{cf_metrics['cash_runway_months']} Bulan" if cf_metrics['cash_runway_months'] < 99 else "Aman (>99 Bulan)"
            metric_card("Cash Runway (Ketahanan)", runway_str, "Berapa lama uang Anda bertahan dengan rata-rata pengeluaran saat ini")

        st.markdown("---")
        c_chart1, c_chart2 = st.columns(2)

        with c_chart1:
            st.subheader("📈 Tren Saldo Berjalan (Running Balance)")
            if col_date and not cf_df.empty:
                fig_balance = px.line(cf_df, x=col_date, y="Running Balance", title="Grafik Saldo Kas dari Waktu ke Waktu", markers=True, color_discrete_sequence=["#7A1F3D"])
                fig_balance.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_balance, use_container_width=True)
            else:
                st.info("Pilih kolom tanggal yang valid untuk melihat grafik tren saldo berjalan.")

        with c_chart2:
            st.subheader("🥧 Proporsi Pengeluaran Berdasarkan Kategori")
            if "Kategori Pengeluaran" in cf_df.columns:
                cat_summary = cf_df[cf_df['Kas Keluar'] > 0].groupby("Kategori Pengeluaran")['Kas Keluar'].sum().reset_index()
                if not cat_summary.empty:
                    fig_pie = px.pie(cat_summary, names="Kategori Pengeluaran", values="Kas Keluar", title="Distribusi Kategori Pengeluaran Kas", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_pie.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("Tidak ada data pengeluaran untuk dikategorikan.")

        st.markdown("---")
        st.subheader("🚨 Deteksi Anomali Pengeluaran Kas")
        if not cf_anomalies.empty:
            st.warning(f"Ditemukan **{len(cf_anomalies)} transaksi pengeluaran** yang nilainya jauh di atas rata-rata normal (potensi pemborosan atau lonjakan biaya tak terduga).")
            st.dataframe(cf_anomalies[[col_date, col_desc, 'Kas Keluar', 'Kategori Pengeluaran']] if col_date and col_desc else cf_anomalies, use_container_width=True)
        else:
            st.success("Tidak ditemukan anomali pengeluaran kas ekstrem.")

        with st.expander("👁️ Lihat Tabel Lengkap Mutasi & Saldo Berjalan (Running Balance)"):
            st.dataframe(cf_df, use_container_width=True)

    # ---------------- TAB 3: ANALISIS RASIO & KEUANGAN ----------------
    with tab3:
        st.subheader(f"Ikhtisar Pos Keuangan ({entity_type})")
        if "Pemerintahan" in entity_type:
            c1, c2, c3, c4 = st.columns(4)
            with c1: metric_card("Realisasi Pendapatan", format_rupiah(fin_values["real_pendapatan"]))
            with c2: metric_card("Realisasi Belanja", format_rupiah(fin_values["real_belanja"]))
            with c3: metric_card("Pendapatan Asli Daerah (PAD)", format_rupiah(fin_values["pad"]))
            with c4: metric_card("Surplus / (Defisit)", format_rupiah(fin_values["surplus"]))
        else:
            c1, c2, c3, c4 = st.columns(4)
            with c1: metric_card("Total Pendapatan", format_rupiah(fin_values["pendapatan"]))
            with c2: metric_card("Laba Bersih", format_rupiah(fin_values["laba"]))
            with c3: metric_card("Total Aset", format_rupiah(fin_values["total_aset"]))
            with c4: metric_card("Total Ekuitas", format_rupiah(fin_values["ekuitas"]))

        st.markdown("---")
        st.subheader("Hasil Perhitungan & Interpretasi Rasio")
        r_cols = st.columns(len(rasio_hasil)) if rasio_hasil else []
        for idx, (r_name, r_val) in enumerate(rasio_hasil.items()):
            with r_cols[idx]:
                metric_card(r_name, f"{r_val}")

        if rasio_hasil:
            fig_rasio = px.bar(
                x=list(rasio_hasil.keys()),
                y=list(rasio_hasil.values()),
                labels={"x": "Indikator Rasio", "y": "Nilai"},
                title="Visualisasi Rasio Keuangan",
                color_discrete_sequence=["#FF5C8B"]
            )
            st.plotly_chart(fig_rasio, use_container_width=True)

            for r_name, r_val in rasio_hasil.items():
                insight_html = get_detailed_insight(r_name, r_val)
                st.markdown(f"""
                <div class="insight-box">
                    <h4>📊 {r_name}: {r_val}</h4>
                    <p>{insight_html}</p>
                </div>
                """, unsafe_allow_html=True)

    # ---------------- TAB 4: ASISTEN AUDITOR AI & WHAT-IF ----------------
    with tab4:
        st.subheader("💬 Konsultasi AI & Simulasi What-If Arus Kas")
        audit_summary_context = json.dumps({
            "file": uploaded_file.name,
            "entitas": entity_type,
            "jumlah_baris": len(df),
            "ringkasan_kas": cf_metrics,
            "temuan_audit": all_findings,
            "pos_keuangan": fin_values,
            "rasio_keuangan": rasio_hasil
        }, ensure_ascii=False, indent=2)

        preset_q = st.selectbox("Pilih Pertanyaan / Simulasi Cepat:", [
            "-- Ketik pertanyaan kustom --",
            "Buatkan ringkasan eksekutif kesehatan keuangan dan likuiditas dari data arus kas ini.",
            "Jika saya memotong biaya operasional sebesar 20%, berapa lama cash runway saya bertambah?",
            "Anomali apa saja yang paling berisiko dalam data transaksi ini dan apa saran auditnya?"
        ])

        user_prompt = st.text_area("Instruksi / Pertanyaan Anda ke AI:", value="" if preset_q == "-- Ketik pertanyaan kustom --" else preset_q, height=100)

        if st.button("🚀 Kirim ke FINORA AI"):
            if user_prompt.strip():
                with st.chat_message("assistant"):
                    st.write_stream(ask_gemini_stream(user_prompt, audit_summary_context, system_role="AI Cash Flow Advisor & Senior Auditor"))
            else:
                st.warning("Silakan masukkan pertanyaan terlebih dahulu.")

    # ---------------- TAB 5: SIMPAN & EKSPOR ----------------
    with tab5:
        st.subheader("💾 Simpan ke Database & Unduh Laporan PDF")
        col_save, col_pdf = st.columns(2)

        with col_save:
            if st.button("Simpan Hasil Analisis ke Riwayat"):
                summary_payload = {
                    "fin_values": fin_values,
                    "rasio": rasio_hasil,
                    "cash_flow_metrics": cf_metrics,
                    "total_findings": len(all_findings)
                }
                aid = save_analysis(file_id, audit_mode, selected_modules, summary_payload)
                save_findings(aid, all_findings)
                st.success(f"Analisis berhasil disimpan ke database dengan ID #{aid}!")

        with col_pdf:
            summary_for_pdf = fin_values if any(fin_values.values()) else {"Total Baris Diperiksa": len(df), "Total Temuan": len(all_findings)}
            pdf_buffer = generate_pdf_report(uploaded_file.name, entity_type, summary_for_pdf, all_findings, rasio_hasil)
            st.download_button(
                label="📄 Unduh Laporan Audit & Kas (PDF)",
                data=pdf_buffer,
                file_name=f"Laporan_FINORA_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf"
            )

else:
    st.info("👆 Silakan unggah file Excel (`.xlsx`, `.xls`) atau `.csv` pada panel di atas untuk memulai analisis audit dan Cash Flow Intelligence.")
