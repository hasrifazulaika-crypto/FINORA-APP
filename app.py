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
.metric-card { background: white; padding: 16px; border-radius: 12px; border-left: 4px solid #FF5C8B; box-shadow: 0 2px 8px rgba(0,0,0,0.05); margin-bottom: 15px; }
.metric-card p { color: #555; font-size: 14px; margin: 0; font-weight: 600; }
.metric-card h3 { color: #7A1F3D; margin: 5px 0 0 0; font-size: 24px; font-weight: 700; }
.insight-box { background-color: white; padding: 20px; border-radius: 10px; border-left: 5px solid #FF5C8B; margin-bottom: 15px; box-shadow: 0 2px 10px rgba(0,0,0,0.03); }
.insight-box h4 { margin-top: 0; color: #7A1F3D; font-size: 18px; margin-bottom: 10px; }
.insight-box p { margin-bottom: 8px; font-size: 14px; line-height: 1.6; color: #333; }
</style>
""", unsafe_allow_html=True)

def format_rupiah(val):
    if abs(val) >= 1_000_000_000_000: return f"Rp {val/1_000_000_000_000:,.2f} T"
    elif abs(val) >= 1_000_000_000: return f"Rp {val/1_000_000_000:,.2f} M"
    elif abs(val) >= 1_000_000: return f"Rp {val/1_000_000:,.2f} Jt"
    else: return f"Rp {val:,.0f}"

def metric_card(label, value, help_text=""):
    st.markdown(f'<div class="metric-card" title="{help_text}"><p>{label}</p><h3>{value}</h3></div>', unsafe_allow_html=True)

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
    df = pd.read_sql_query("""SELECT a.id, a.summary_json, f.filename FROM analyses a JOIN files f ON a.file_id = f.id WHERE a.id=?""", conn, params=(aid,))
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
        model_id = "gemini-3.8-flash"
        full_prompt = f"Kamu adalah FINORA AI, {system_role}.\nJawab berdasarkan konteks data secara ringkas dan profesional.\n\n[KONTEKS]\n{audit_context}\n\n[PERTANYAAN]\n{prompt}"
        
        if hasattr(client, "interactions"):
            try:
                stream = client.interactions.create(model=model_id, input=full_prompt, stream=True)
                for event in stream:
                    delta = getattr(event, "delta", None)
                    if delta and getattr(delta, "text", None): yield delta.text
                    elif getattr(event, "text", None): yield event.text
                return
            except Exception: pass 
        
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
        if not mapping["debit"] and any(k in c for k in ["debit", "masuk", "pemasukan", "in"]): mapping["debit"] = col
        if not mapping["credit"] and any(k in c for k in ["kredit", "keluar", "pengeluaran", "out"]): mapping["credit"] = col
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
            findings.append({"module": "Kelengkapan Data", "risk_level": "Tinggi", "title": f"{len(flagged)} baris data vital kosong", "detail": "Data kosong pada kolom vital.", "recommendation": "Lengkapi data."})
    return findings, flagged

def audit_duplicates(df, cols_to_check):
    findings, flagged = [], pd.DataFrame()
    if cols_to_check:
        flagged = df[df.duplicated(subset=cols_to_check, keep=False)]
        if not flagged.empty:
            findings.append({"module": "Data Duplikat", "risk_level": "Tinggi", "title": f"{len(flagged)} baris ganda", "detail": "Transaksi sama persis.", "recommendation": "Cek double-entry."})
    return findings, flagged

def audit_negative(df, amount_col):
    findings, flagged = [], pd.DataFrame()
    if amount_col:
        s = pd.to_numeric(df[amount_col], errors="coerce")
        flagged = df[s < 0]
        if not flagged.empty:
            findings.append({"module": "Nilai Negatif", "risk_level": "Rendah", "title": f"{len(flagged)} nilai negatif", "detail": "Ditemukan nilai minus.", "recommendation": "Uji petik."})
    return findings, flagged

def audit_overbudget(df, budget_col, amount_col):
    findings, flagged = [], pd.DataFrame()
    if budget_col and amount_col:
        try:
            b = pd.to_numeric(df[budget_col], errors="coerce").fillna(0)
            a = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)
            flagged = df[a > b]
            if not flagged.empty:
                findings.append({"module": "Overbudget", "risk_level": "Tinggi", "title": f"{len(flagged)} akun melebihi pagu", "detail": "Realisasi > Anggaran.", "recommendation": "Minta penjelasan."})
        except: pass
    return findings, flagged

def audit_weekend(df, date_col):
    findings, flagged = [], pd.DataFrame()
    if date_col:
        try:
            dates = pd.to_datetime(df[date_col], errors="coerce")
            flagged = df[dates.dt.dayofweek >= 5]
            if not flagged.empty:
                findings.append({"module": "Transaksi Weekend", "risk_level": "Tinggi", "title": f"{len(flagged)} transaksi weekend", "detail": "Transaksi di luar hari kerja.", "recommendation": "Periksa otorisasi."})
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
                findings.append({"module": "Kata Kunci Manipulasi", "risk_level": "Sedang", "title": f"{len(flagged)} transaksi mencurigakan", "detail": "Mengandung kata adjustment/void.", "recommendation": "Minta penjelasan."})
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
                    findings.append({"module": "Anomali Nominal", "risk_level": "Tinggi", "title": f"{len(flagged)} nilai ekstrem (Outlier)", "detail": f"Batas wajar: {lower:,.0f} - {upper:,.0f}", "recommendation": "Uji petik transaksi."})
        except: pass
    return findings, flagged

# ============================================================
# FINANCIAL & CASHFLOW ENGINE
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
        results["real_pendapatan"] = get_total(real_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan daerah'], ['pendapatan asli daerah', 'dana perimbangan'])
        results["ang_pendapatan"] = get_total(ang_col, ['total pendapatan', 'jumlah pendapatan', 'pendapatan daerah'], ['pendapatan asli daerah', 'dana perimbangan'])
        results["real_belanja"] = get_total(real_col, ['total belanja', 'jumlah belanja', 'belanja daerah'], ['belanja operasi', 'belanja modal'])
        results["ang_belanja"] = get_total(ang_col, ['total belanja', 'jumlah belanja', 'belanja daerah'], ['belanja operasi', 'belanja modal'])
        results["pad"] = get_total(real_col, ['pendapatan asli daerah', 'pad'], ['pendapatan asli daerah'])
        results["surplus"] = results["real_pendapatan"] - results["real_belanja"]
    else:
        results["pendapatan"] = get_total(real_col, ['total pendapatan', 'pendapatan', 'penjualan'], ['pendapatan', 'penjualan', 'revenue'])
        results["laba"] = get_total(real_col, ['laba bersih', 'laba tahun berjalan'], ['laba bersih', 'net income'])
        results["total_aset"] = get_total(real_col, ['total aset', 'jumlah aset'], ['total aset'])
        results["ekuitas"] = get_total(real_col, ['total ekuitas', 'total modal'], ['ekuitas'])
        results["total_liabilitas"] = get_total(real_col, ['total liabilitas', 'total utang'], ['liabilitas', 'utang'])
        results["aset_lancar"] = get_total(real_col, ['total aset lancar'], ['aset lancar'])
        results["liabilitas_lancar"] = get_total(real_col, ['total liabilitas jangka pendek'], ['liabilitas jangka pendek'])
    return results

def hitung_rasio_dari_nilai(nilai, entity_type):
    rasio = {}
    if "Pemerintahan" in entity_type:
        rasio["Efektivitas Pendapatan (%)"] = round((nilai.get("real_pendapatan", 0) / nilai.get("ang_pendapatan", 1)) * 100, 2) if nilai.get("ang_pendapatan", 0) > 0 else 0.0
        rasio["Efisiensi Belanja (%)"] = round((nilai.get("real_belanja", 0) / nilai.get("ang_belanja", 1)) * 100, 2) if nilai.get("ang_belanja", 0) > 0 else 0.0
        rasio["Kemandirian Daerah (%)"] = round((nilai.get("pad", 0) / nilai.get("real_pendapatan", 1)) * 100, 2) if nilai.get("real_pendapatan", 0) > 0 else 0.0
        rasio["Pertumbuhan Pendapatan (%)"] = 0.0 
    else:
        rasio["Current Ratio (x)"] = round(nilai.get("aset_lancar", 0) / nilai.get("liabilitas_lancar", 1), 2) if nilai.get("liabilitas_lancar", 0) > 0 else 0.0
        rasio["Net Profit Margin (%)"] = round((nilai.get("laba", 0) / nilai.get("pendapatan", 1)) * 100, 2) if nilai.get("pendapatan", 0) > 0 else 0.0
        rasio["ROA (%)"] = round((nilai.get("laba", 0) / nilai.get("total_aset", 1)) * 100, 2) if nilai.get("total_aset", 0) > 0 else 0.0
        rasio["ROE (%)"] = round((nilai.get("laba", 0) / nilai.get("ekuitas", 1)) * 100, 2) if nilai.get("ekuitas", 0) > 0 else 0.0
        rasio["Debt to Equity (x)"] = round(nilai.get("total_liabilitas", 0) / nilai.get("ekuitas", 1), 2) if nilai.get("ekuitas", 0) > 0 else 0.0
    return rasio

def get_detailed_insight(ratio_name, val):
    if val == 0 or pd.isna(val): return "<b>Kondisi:</b> Data 0/Kosong. Periksa input file Anda."
    if "Efektivitas" in ratio_name:
        if val >= 100: return "<b>Kondisi:</b> Sangat Baik (Realisasi > Target).<br><b>Rekomendasi:</b> Pertahankan kinerja."
        elif val >= 90: return "<b>Kondisi:</b> Baik.<br><b>Rekomendasi:</b> Tingkatkan kepatuhan."
        else: return "<b>Kondisi:</b> Kurang.<br><b>Rekomendasi:</b> Lakukan intensifikasi & ekstensifikasi pajak."
    elif "Efisiensi" in ratio_name:
        if val > 100: return "<b>Kondisi:</b> Overbudget.<br><b>Rekomendasi:</b> Evaluasi perencanaan anggaran."
        elif val >= 90: return "<b>Kondisi:</b> Efisien.<br><b>Rekomendasi:</b> Pastikan output berkualitas."
        else: return "<b>Kondisi:</b> Penyerapan Rendah.<br><b>Rekomendasi:</b> Percepat proses tender."
    elif "Kemandirian" in ratio_name:
        if val >= 50: return "<b>Kondisi:</b> Mandiri.<br><b>Rekomendasi:</b> Kembangkan BUMD."
        elif val >= 25: return "<b>Kondisi:</b> Sedang.<br><b>Rekomendasi:</b> Gali potensi PAD baru."
        else: return "<b>Kondisi:</b> Rendah.<br><b>Rekomendasi:</b> Reformasi perpajakan daerah."
    elif "Current Ratio" in ratio_name:
        if val >= 1.5: return "<b>Kondisi:</b> Sehat.<br><b>Rekomendasi:</b> Manfaatkan kas untuk investasi."
        else: return "<b>Kondisi:</b> Risiko Likuiditas.<br><b>Rekomendasi:</b> Percepat penagihan piutang."
    return "<b>Analisis:</b> Detail belum tersedia."

# --- FITUR BARU: CASHFLOW ENGINE ---
def auto_extract_cashflow(df):
    results = {"operasi": 0.0, "investasi": 0.0, "pendanaan": 0.0, "net_cash": 0.0, "kas_awal": 0.0, "kas_akhir": 0.0}
    
    str_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
    num_cols = df.select_dtypes(include=['number', 'float', 'int']).columns.tolist()
    if not str_cols or not num_cols: return results
    name_col = str_cols[0] 
    val_col = num_cols[0]

    clean_names = df[name_col].astype(str).str.lower().str.replace(r'[^a-z\s]', '', regex=True).str.strip()

    def get_total(contains_keys):
        mask = clean_names.str.contains('|'.join(contains_keys), na=False)
        if mask.any():
            return pd.to_numeric(df.loc[mask, val_col], errors='coerce').fillna(0).sum()
        return 0.0

    results["operasi"] = get_total(['arus kas operasi', 'kas dari operasi', 'operating activities', 'operasi bersih', 'net operating'])
    results["investasi"] = get_total(['arus kas investasi', 'kas untuk investasi', 'investing activities', 'pembelian aset tetap', 'net investing'])
    results["pendanaan"] = get_total(['arus kas pendanaan', 'kas dari pendanaan', 'financing activities', 'pinjaman', 'modal disetor', 'net financing'])
    results["net_cash"] = get_total(['kenaikan kas', 'penurunan kas', 'net cash', 'arus kas bersih', 'perubahan kas'])
    results["kas_awal"] = get_total(['kas awal', 'saldo awal kas', 'beginning cash'])
    results["kas_akhir"] = get_total(['kas akhir', 'saldo akhir kas', 'ending cash'])

    if results["net_cash"] == 0:
        results["net_cash"] = results["operasi"] + results["investasi"] + results["pendanaan"]

    return results

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
    summary_data = [["No", "Kode", "Kategori Uji", "Risiko", "Deskripsi Singkat"]]
    for idx, f in enumerate(findings_list, 1):
        summary_data.append([str(idx), f"F-{idx:02d}", f['module'], f['risk_level'], f['title']])
    t_summary = Table(summary_data, colWidths=[30, 50, 100, 60, 230])
    t_summary.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF5C8B')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 9), ('GRID', (0, 0), (-1, -1), 1, colors.grey)]))
    elements.append(t_summary)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph("B. Rincian Temuan & Bukti", heading_style))
    for idx, f in enumerate(findings_list, 1):
        elements.append(Paragraph(f"<b>F-{idx:02d}: {f['title']}</b>", normal_style))
        elements.append(Paragraph(f"<b>1. Kondisi:</b> {f['detail']}", normal_style))
        elements.append(Paragraph(f"<b>2. Rekomendasi:</b> {f['recommendation']}", normal_style))
        modul_name = f['module']
        df_flag = pd.DataFrame()
        if modul_name in flagged_data and not flagged_data[modul_name].empty: df_flag = flagged_data[modul_name]
        elif "Kelengkapan" in modul_name and "Kelengkapan Kosong" in flagged_data: df_flag = flagged_data["Kelengkapan Kosong"]
        elif "Duplikat" in modul_name and "Duplikasi Ekstrem" in flagged_data: df_flag = flagged_data["Duplikasi Ekstrem"]
        elif "Negatif" in modul_name and "Nilai Negatif" in flagged_data: df_flag = flagged_data["Nilai Negatif"]
        elif "Overbudget" in modul_name and "Overbudget" in flagged_data: df_flag = flagged_data["Overbudget"]
        if not df_flag.empty:
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
        if "Pemerintahan" in entity_type:
            data_nilai = [{"Komponen": "Realisasi Pendapatan", "Nilai": nilai.get("real_pendapatan", 0)}, {"Komponen": "Anggaran Pendapatan", "Nilai": nilai.get("ang_pendapatan", 0)}, {"Komponen": "Realisasi Belanja", "Nilai": nilai.get("real_belanja", 0)}, {"Komponen": "Anggaran Belanja", "Nilai": nilai.get("ang_belanja", 0)}]
        else:
            data_nilai = [{"Komponen": "Total Pendapatan", "Nilai": nilai.get("pendapatan", 0)}, {"Komponen": "Laba Bersih", "Nilai": nilai.get("laba", 0)}, {"Komponen": "Total Aset", "Nilai": nilai.get("total_aset", 0)}]
        pd.DataFrame(data_nilai).to_excel(writer, sheet_name="Data Keuangan", index=False)
        pd.DataFrame([{"Rasio": k, "Nilai (%)": v} for k, v in rasio.items()]).to_excel(writer, sheet_name="Rasio Keuangan
