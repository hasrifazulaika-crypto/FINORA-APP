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
    api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    if not api_key: return None
    return genai.Client(api_key=api_key)

def ask_gemini_stream(prompt, audit_context, system_role="Asisten Auditor Profesional"):
    try:
        client = get_gemini_client()
        if client is None:
            yield "⚠️ API Key Gemini belum diatur di Secrets."
            return
        model_id = "gemini-2.5-flash"
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
            if getattr(chunk, "text", None): yield chunk.text
    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan: {str(e)}"

# ============================================================
# AUDIT ENGINE V5
# ============================================================
def auto_detect_columns(df):
    mapping = {"date": None, "amount": None, "desc": None, "debit": None, "credit": None, "budget": None, "invoice": None}
    for col in df.columns:
        c = str(col).lower()
        if not mapping["date"] and any(k in c for k in ["tanggal", "tgl", "date", "waktu", "periode"]): mapping["date"] = col
        if not mapping["amount"] and any(k in c for k in ["actual", "realisasi", "nominal", "jumlah", "total", "nilai", "amount", "harga"]): mapping["amount"] = col
        if not mapping["budget"] and any(k in c for k in ["anggaran", "budget", "pagu"]): mapping["budget"] = col
        if not mapping["desc"] and any(k in c for k in ["deskripsi", "keterangan", "uraian", "category", "kategori", "nama", "item", "barang"]): mapping["desc"] = col
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
            findings.append({"module": "Data Duplikat", "risk_level": "Tinggi", "title": f"{len(flagged)} baris transaksi ganda", "detail": "Terdapat transaksi dengan nilai/deskripsi yang sama persis.", "recommendation": "Pastikan bukan input dua kali."})
    return findings, flagged

def audit_negative(df, amount_col):
    findings, flagged = [], pd.DataFrame()
    if amount_col:
        s = pd.to_numeric(df[amount_col], errors="coerce")
        flagged = df[s < 0]
        if not flagged.empty:
            findings.append({"module": "Nilai Negatif", "risk_level": "Rendah", "title": f"{len(flagged)} transaksi bernilai negatif", "detail": "Ditemukan nilai minus pada kolom nominal.", "recommendation": "Pastikan ada memo retur/koreksi."})
    return findings, flagged

def audit_overbudget(df, budget_col, amount_col):
    findings, flagged = [], pd.DataFrame()
    if budget_col and amount_col:
        try:
            b = pd.to_numeric(df[budget_col], errors="coerce").fillna(0)
            a = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
            flagged = df[a > b]
            if not flagged.empty:
                findings.append({"module": "Overbudget", "risk_level": "Tinggi", "title": f"{len(flagged)} akun melebihi pagu", "detail": "Realisasi melampaui anggaran.", "recommendation": "Evaluasi pergeseran anggaran."})
        except: pass
    return findings, flagged

def audit_weekend(df, date_col):
    findings, flagged = [], pd.DataFrame()
    if date_col:
        try:
            dates = pd.to_datetime(df[date_col], errors="coerce")
            flagged = df[dates.dt.dayofweek >= 5]
            if not flagged.empty:
                findings.append({"module": "Transaksi Weekend", "risk_level": "Tinggi", "title": f"{len(flagged)} transaksi akhir pekan", "detail": "Transaksi di luar hari kerja.", "recommendation": "Periksa otorisasi transaksi."})
        except: pass
    return findings, flagged

def audit_suspicious_words(df, desc_col):
    findings, flagged = [], pd.DataFrame()
    keywords = ["adjustment", "adj", "void", "koreksi", "manual", "write-off"]
    if desc_col:
        try:
            pattern = '|'.join(keywords)
            flagged = df[df[desc_col].astype(str).str.lower().str.contains(pattern, na=False)]
            if not flagged.empty:
                findings.append({"module": "Kata Kunci Mencurigakan", "risk_level": "Sedang", "title": f"{len(flagged)} transaksi adjustment/void", "detail": "Mengandung kata kunci penyesuaian manual.", "recommendation": "Minta penjelasan tertulis."})
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
                    findings.append({"module": "Outlier IQR", "risk_level": "Tinggi", "title": f"{len(flagged)} nilai ekstrem", "detail": f"Di luar batas wajar: {lower:,.0f} s.d {upper:,.0f}", "recommendation": "Uji petik transaksi ekstrem."})
        except: pass
    return findings, flagged

# ============================================================
# SMART FINANCIAL ENGINE
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
            m_ex = clean_names.isin(exact_keys)
            if m_ex.any(): return pd.to_numeric(df.loc[m_ex, target_col], errors='coerce').fillna(0).sum()
        if contains_keys:
            m_co = clean_names.str.contains('|'.join(contains_keys), na=False)
            if m_co.any(): return pd.to_numeric(df.loc[m_co, target_col], errors='coerce').fillna(0).sum()
        return 0.0

    if "Pemerintahan" in entity_type:
        real_p = get_total(real_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan'])
        if real_p == 0: real_p = get_total(real_col, [], ['pendapatan asli daerah', 'dana transfer'])
        ang_p = get_total(ang_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan'])
        real_b = get_total(real_col, ['total belanja', 'jumlah belanja', 'belanja'])
        if real_b == 0: real_b = get_total(real_col, [], ['belanja operasi', 'belanja modal'])
        ang_b = get_total(ang_col, ['total belanja', 'jumlah belanja', 'belanja'])
        pad = get_total(real_col, ['pendapatan asli daerah', 'pad'], ['pendapatan asli daerah'])
        surplus = real_p - real_b
        results.update({"real_pendapatan": real_p, "ang_pendapatan": ang_p, "real_belanja": real_b, "ang_belanja": ang_b, "pad": pad, "surplus": surplus})
    else:
        results["pendapatan"] = get_total(real_col, ['pendapatan', 'penjualan'], ['pendapatan', 'penjualan', 'revenue'])
        results["laba"] = get_total(real_col, ['laba bersih', 'laba tahun berjalan'], ['laba bersih', 'net income'])
        results["total_aset"] = get_total(real_col, ['total aset', 'jumlah aset'], ['total aset'])
        results["ekuitas"] = get_total(real_col, ['total ekuitas', 'total modal'], ['ekuitas'])
        results["total_liabilitas"] = get_total(real_col, ['total liabilitas', 'total utang'], ['liabilitas'])
        results["aset_lancar"] = get_total(real_col, ['total aset lancar'], ['aset lancar'])
        results["liabilitas_lancar"] = get_total(real_col, ['total liabilitas jangka pendek'], ['liabilitas jangka pendek'])
    return results

def hitung_rasio_dari_nilai(nilai, entity_type):
    rasio = {}
    if "Pemerintahan" in entity_type:
        real_p, ang_p = nilai.get("real_pendapatan", 0), nilai.get("ang_pendapatan", 0)
        real_b, ang_b = nilai.get("real_belanja", 0), nilai.get("ang_belanja", 0)
        pad = nilai.get("pad", 0)
        rasio["Efektivitas Pendapatan (%)"] = round((real_p / ang_p) * 100, 2) if ang_p > 0 else 0.0
        rasio["Efisiensi Belanja (%)"] = round((real_b / ang_b) * 100, 2) if ang_b > 0 else 0.0
        rasio["Kemandirian Daerah (%)"] = round((pad / real_p) * 100, 2) if real_p > 0 else 0.0
    else:
        pendapatan, laba = nilai.get("pendapatan", 0), nilai.get("laba", 0)
        aset_lancar, liab_lancar = nilai.get("aset_lancar", 0), nilai.get("liabilitas_lancar", 0)
        total_aset, total_liab, ekuitas = nilai.get("total_aset", 0), nilai.get("total_liabilitas", 0), nilai.get("ekuitas", 0)
        rasio["Current Ratio (x)"] = round(aset_lancar / liab_lancar, 2) if liab_lancar > 0 else 0.0
        rasio["Net Profit Margin (%)"] = round((laba / pendapatan) * 100, 2) if pendapatan > 0 else 0.0
        rasio["ROA (%)"] = round((laba / total_aset) * 100, 2) if total_aset > 0 else 0.0
        rasio["ROE (%)"] = round((laba / ekuitas) * 100, 2) if ekuitas > 0 else 0.0
        rasio["Debt to Equity (x)"] = round(total_liab / ekuitas, 2) if ekuitas > 0 else 0.0
    return rasio

def get_detailed_insight(ratio_name, val):
    if val == 0:
        return "<b>Kondisi:</b> 0 (Data belum terisi/terbaca).<br><b>Rekomendasi:</b> Pastikan struktur kolom file Excel Anda sudah sesuai."
    return f"<b>Kondisi Nilai:</b> {val}. Analisis berjalan normal berdasarkan input data terkini."

# ============================================================
# EXPORT FUNCTIONS
# ============================================================
def generate_excel(flagged_data_dict, findings_list):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_findings = pd.DataFrame(findings_list)
        if not df_findings.empty:
            df_findings.columns = ["Modul", "Tingkat Risiko", "Judul", "Detail", "Rekomendasi"]
            df_findings.to_excel(writer, sheet_name="Ringkasan Temuan", index=False)
        for modul, df_flag in flagged_data_dict.items():
            if not df_flag.empty:
                df_flag.to_excel(writer, sheet_name=modul[:31], index=False)
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

    header_data = [
        ["Nama Entitas", f": {entitas if entitas else '-'}"],
        ["Periode Audit", f": {periode if periode else '-'}"],
        ["Nama File", f": {filename}"],
        ["Tanggal Pemeriksaan", f": {datetime.now().strftime('%d %B %Y')}"],
        ["Skor Risiko", f": {score}/100"]
    ]
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
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF5C8B')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
    ]))
    elements.append(t_summary)
    elements.append(Spacer(1, 20))

    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_financial_excel(nilai, rasio):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        pd.DataFrame(list(nilai.items()), columns=["Komponen", "Nilai"]).to_excel(writer, sheet_name="Data Keuangan", index=False)
        pd.DataFrame(list(rasio.items()), columns=["Rasio", "Nilai"]).to_excel(writer, sheet_name="Rasio Keuangan", index=False)
    buffer.seek(0)
    return buffer

def generate_financial_pdf(nilai, rasio, entity_type, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='Title', parent=styles['Heading1'], textColor=colors.HexColor('#7A1F3D'), alignment=1)
    heading_style = ParagraphStyle(name='Heading', parent=styles['Heading2'], textColor=colors.HexColor('#FF5C8B'))
    normal_style = styles['Normal']

    elements.append(Paragraph("LAPORAN ANALISIS KEUANGAN", title_style))
    elements.append(Paragraph("FINORA Financial Intelligence System", normal_style))
    elements.append(Spacer(1, 10))

    header_data = [
        ["Jenis Entitas", f": {entity_type}"],
        ["Nama File", f": {filename}"],
        ["Tanggal Analisis", f": {datetime.now().strftime('%d %B %Y')}"]
    ]
    t_header = Table(header_data, colWidths=[120, 350])
    t_header.setStyle(TableStyle([('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 10)]))
    elements.append(t_header)
    elements.append(Spacer(1, 20))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# UI COMPONENTS & KKA RENDER
# ============================================================
def finora_header(title, subtitle):
    logo_html = f'<img src="data:image/png;base64,{LOGO_B64}" style="height: 40px; vertical-align: middle; margin-right: 12px; margin-top: -5px; border-radius: 8px;">' if LOGO_B64 else ''
    st.markdown(f'<div class="finora-header"><h1>{logo_html} {title}</h1><p>{subtitle}</p></div>', unsafe_allow_html=True)

def render_audit_results(all_findings, flagged_data, filename, entitas, periode, score):
    color = "#C62828" if score > 50 else "#EF6C00" if score > 20 else "#2E7D32"
    st.markdown(f"""
    <div style="background-color: white; padding: 20px; border-radius: 10px; border: 2px solid #FF5C8B; margin-bottom: 20px;">
        <h2 style="color: #7A1F3D; text-align: center; margin-bottom: 5px;">KERTAS KERJA AUDIT (KKA)</h2>
        <p style="text-align: center; color: gray; margin-top: 0;">FINORA Financial Intelligence System</p>
        <hr>
        <table style="width: 100%; font-size: 14px;">
            <tr><td><b>Nama Entitas</b></td><td>: {entitas if entitas else '-'}</td></tr>
            <tr><td><b>Periode Audit</b></td><td>: {periode if periode else '-'}</td></tr>
            <tr><td><b>Nama File</b></td><td>: {filename}</td></tr>
            <tr><td><b>Skor Risiko</b></td><td>: <span style='color:{color}; font-weight:bold; font-size:18px;'>{score}/100</span></td></tr>
        </table>
    </div>
    """, unsafe_allow_html=True)

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
                st.markdown(f"**Kondisi:** {f['detail']}")
                st.error(f"**Rekomendasi:** {f['recommendation']}")
                modul_name = f['module']
                if modul_name in flagged_data and not flagged_data[modul_name].empty:
                    st.dataframe(flagged_data[modul_name].head(10), use_container_width=True)

        st.markdown("---")
        col1, col2 = st.columns(2)
        excel_buffer = generate_excel(flagged_data, all_findings)
        col1.download_button("⬇️ Download Excel", data=excel_buffer, file_name=f"KKA_{entitas}.xlsx", use_container_width=True)
        pdf_buffer = generate_pdf(entitas, periode, filename, score, all_findings, flagged_data)
        col2.download_button("⬇️ Download PDF", data=pdf_buffer, file_name=f"KKA_{entitas}.pdf", use_container_width=True)

    with tab_grafik:
        if all_findings:
            risk_df = pd.DataFrame(all_findings)["risk_level"].value_counts().reset_index()
            fig_pie = px.pie(risk_df, names="risk_level", values="count", color="risk_level", color_discrete_map={"Tinggi": "#C62828", "Sedang": "#EF6C00", "Rendah": "#2E7D32"}, hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
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

# ============================================================
# HALAMAN MODUL
# ============================================================
def page_audit():
    finora_header("Audit Profesional", "Sistem deteksi otomatis anomali transaksi")
    st.markdown("### 1. Upload Data Audit")
    c1, c2 = st.columns([3, 1])
    uploaded = c1.file_uploader("Upload File Excel/CSV", type=["xlsx", "csv"], key="audit_upload_unique")
    header_row = c2.number_input("Header di baris ke-", min_value=0, value=0, key="audit_header_unique")

    if uploaded:
        df = pd.read_excel(uploaded, header=header_row) if uploaded.name.endswith('xlsx') else pd.read_csv(uploaded, header=header_row)
        detected = auto_detect_columns(df)
        
        st.markdown("### 2. Pemetaan Kolom")
        with st.expander("⚙️ Pengaturan Kolom", expanded=False):
            c1, c2 = st.columns(2)
            date_col = c1.selectbox("Kolom Tanggal", [""] + list(df.columns), index=list(df.columns).index(detected["date"])+1 if detected["date"] else 0, key="audit_col_date")
            desc_col = c2.selectbox("Kolom Deskripsi", [""] + list(df.columns), index=list(df.columns).index(detected["desc"])+1 if detected["desc"] else 0, key="audit_col_desc")
            c3, c4, c5 = st.columns(3)
            amount_col = c3.selectbox("Kolom Nominal", [""] + list(df.columns), index=list(df.columns).index(detected["amount"])+1 if detected["amount"] else 0, key="audit_col_amt")
            budget_col = c4.selectbox("Kolom Anggaran", [""] + list(df.columns), index=list(df
