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

def fetch_analysis_by_id(aid):
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT a.id, a.summary_json, f.filename 
        FROM analyses a JOIN files f ON a.file_id = f.id 
        WHERE a.id=?
    """, conn, params=(aid,))
    conn.close(); return df.iloc[0] if not df.empty else None

def fetch_findings(aid):
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM findings WHERE analysis_id=?", conn, params=(aid,))
    conn.close(); return df

# ============================================================
# AI ENGINE (FIXED MODEL VERSION)
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

        # Menggunakan model stabil yang didukung penuh
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
            if getattr(chunk, "text", None): 
                yield chunk.text

    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan koneksi AI: {str(e)}"

# ============================================================
# AUDIT ENGINE & CASH FLOW ENGINE
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

def process_cash_flow_data(df, date_col, desc_col, in_col, out_col, initial_balance=0.0):
    pdf = df.copy()
    if date_col:
        pdf[date_col] = pd.to_datetime(pdf[date_col], errors="coerce")
        pdf = pdf.sort_values(by=date_col).reset_index(drop=True)
        pdf['Bulan'] = pdf[date_col].dt.to_period('M').astype(str)
    else:
        pdf['Bulan'] = 'Tidak Diketahui'

    pdf['Kas Masuk'] = pd.to_numeric(pdf[in_col], errors="coerce").fillna(0.0) if in_col else 0.0
    pdf['Kas Keluar'] = pd.to_numeric(pdf[out_col], errors="coerce").fillna(0.0) if out_col else 0.0
    pdf['Net Cash Flow'] = pdf['Kas Masuk'] - pdf['Kas Keluar']
    pdf['Running Balance'] = initial_balance + pdf['Net Cash Flow'].cumsum()

    def auto_categorize(text):
        t = str(text).lower()
        if any(k in t for k in ["bensin", "bbm", "tol", "parkir", "transport", "travel"]): return "Transportasi & Perjalanan"
        elif any(k in t for k in ["gaji", "honor", "tunjangan", "payroll", "salary"]): return "Gaji & SDM"
        elif any(k in t for k in ["listrik", "air", "internet", "wifi", "utility", "pln"]): return "Utilitas & Kantor"
        elif any(k in t for k in ["sewa", "rent", "gedung"]): return "Sewa & Fasilitas"
        elif any(k in t for k in ["makan", "konsumsi", "snack", "f&b"]): return "Konsumsi"
        elif any(k in t for k in ["pajak", "tax"]): return "Pajak & Legal"
        elif any(k in t for k in ["penjualan", "sales", "pendapatan", "revenue", "terima"]): return "Penerimaan / Penjualan"
        else: return "Operasional Lainnya"

    if desc_col:
        pdf['Kategori Pengeluaran'] = pdf[desc_col].apply(auto_categorize)
    else:
        pdf['Kategori Pengeluaran'] = "Umum"

    total_out = pdf['Kas Keluar'].sum()
    months_count = 1.0
    if date_col and not pdf[date_col].dropna().empty:
        diff_days = (pdf[date_col].max() - pdf[date_col].min()).days
        months_count = max(1, diff_days / 30.44)

    monthly_burn_rate = total_out / months_count
    current_saldo = pdf['Running Balance'].iloc[-1] if not pdf.empty else initial_balance
    cash_runway_months = round(current_saldo / monthly_burn_rate, 1) if monthly_burn_rate > 0 else 99.9

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

def auto_extract_financials(df, entity_type):
    results = {
        "pendapatan": 0.0, "laba": 0.0, "aset_lancar": 0.0, "liabilitas_lancar": 0.0,
        "total_aset": 0.0, "total_liabilitas": 0.0, "ekuitas": 0.0,
        "real_pendapatan": 0.0, "ang_pendapatan": 0.0,
        "real_belanja": 0.0, "ang_belanja": 0.0, "surplus": 0.0, "pad": 0.0
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
        real_p = get_total(real_col, ['total pendapatan', 'pendapatan daerah', 'pendapatan'])
        ang_p = get_total(ang_col, ['total pendapatan', 'pendapatan daerah', 'pendapatan'])
        real_b = get_total(real_col, ['total belanja', 'belanja daerah', 'belanja'])
        ang_b = get_total(ang_col, ['total belanja', 'belanja daerah', 'belanja'])
        pad = get_total(real_col, ['pendapatan asli daerah', 'pad'], ['pad'])
        surplus = real_p - real_b
        results.update({"real_pendapatan": real_p, "ang_pendapatan": ang_p, "real_belanja": real_b, "ang_belanja": ang_b, "pad": pad, "surplus": surplus})
    else:
        results["pendapatan"] = get_total(real_col, ['total pendapatan', 'pendapatan', 'penjualan'], ['pendapatan', 'penjualan'])
        results["laba"] = get_total(real_col, ['laba bersih', 'laba tahun berjalan'], ['laba bersih'])
        results["total_aset"] = get_total(real_col, ['total aset', 'jumlah aset'], ['total aset'])
        results["ekuitas"] = get_total(real_col, ['total ekuitas', 'total modal'], ['ekuitas'])
        results["total_liabilitas"] = get_total(real_col, ['total liabilitas', 'total utang'], ['liabilitas'])
        results["aset_lancar"] = get_total(real_col, ['total aset lancar'], ['aset lancar'])
        results["liabilitas_lancar"] = get_total(real_col, ['total liabilitas jangka pendek'], ['liabilitas jangka pendek'])
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
    return f"<b>Kondisi:</b> Indikator {ratio_name} tercatat sebesar {val}. Evaluasi secara berkala untuk menjaga stabilitas keuangan."

def generate_excel(flagged_data_dict, findings_list):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_findings = pd.DataFrame(findings_list)
        if not df_findings.empty:
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
        ["Nama Entitas", f": {entitas or '-'}"],
        ["Periode Audit", f": {periode or '-'}"],
        ["Nama File", f": {filename}"],
        ["Skor Risiko", f": {score}/100"]
    ]
    t_header = Table(header_data, colWidths=[120, 350])
    elements.append(t_header)
    elements.append(Spacer(1, 20))
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# KONTROL NAVIGASI SESUAI PERMINTAAN ANDA
# ============================================================
if 'app_started' not in st.session_state:
    st.session_state.app_started = False

if not st.session_state.app_started:
    # --- 1. TAMPILAN AWAL (LANDING PAGE) ---
    st.markdown("""<style>[data-testid="collapsedControl"], [data-testid="stSidebar"] { display: none; }</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists(LOGO_FILENAME):
            st.image(LOGO_FILENAME, width=250)
        else:
            st.markdown("<h1 style='text-align:center; font-size: 5rem;'>F</h1>", unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #FFF0F4 0%, #FFF5E8 100%); border-radius: 20px; margin-bottom: 40px; border: 2px solid #FF5C8B; box-shadow: 0 10px 30px rgba(255, 92, 139, 0.15);">
        <h1 style="color: #7A1F3D; font-size: 3.5rem; margin-bottom: 15px; font-weight: 800;">FINORA AI</h1>
        <p style="color: #2D1B24; font-size: 1.3rem; max-width: 800px; margin: 0 auto; line-height: 1.6;">
            Platform Kecerdasan Finansial untuk mendeteksi anomali, analisis rasio laporan keuangan, serta monitoring <b>Cash Flow Intelligence</b> secara otomatis.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_cta1, col_cta2, col_cta3 = st.columns([1, 2, 1])
    with col_cta2:
        st.markdown("""<style>.stButton > button { background: linear-gradient(135deg, #FF5C8B 0%, #7A1F3D 100%) !important; color: white !important; font-size: 1.3rem !important; padding: 20px 30px !important; border-radius: 12px !important; border: none !important; font-weight: 800 !important; width: 100%; }</style>""", unsafe_allow_html=True)
        if st.button("🚀 MULAI APLIKASI SEKARANG", use_container_width=True):
            st.session_state.app_started = True
            st.rerun()

else:
    # --- 2. SETELAH MASUK: SIDEBAR & FITUR UTAMA ---
    st.markdown("""<style>[data-testid="collapsedControl"] { display: block; }</style>""", unsafe_allow_html=True)
    
    with st.sidebar:
        st.header("🧭 Navigasi Menu")
        menu = st.radio("Pilih Halaman:", [
            "🏠 Dashboard Ringkasan", 
            "🔎 Fitur 1: Audit Profesional", 
            "📈 Fitur 2: Analisis Laporan Keuangan", 
            "💵 Fitur 3: Cash Flow Intelligence",
            "💾 Arsip & Riwayat Laporan"
        ])
        
        st.markdown("---")
        entity_type = st.selectbox("Jenis Entitas", ["Perusahaan / UMKM (Komersial)", "Instansi Pemerintahan / Sektor Publik"])
        
        if st.button("🚪 Kembali ke Halaman Utama"):
            st.session_state.app_started = False
            st.rerun()

    # ---------------- 🏠 DASHBOARD ----------------
    if menu == "🏠 Dashboard Ringkasan":
        st.markdown("""
        <div class="finora-header">
            <h1>📊 Dashboard Utama FINORA AI</h1>
            <p>Pusat kendali dan ringkasan aktivitas analisis data keuangan Anda.</p>
        </div>
        """, unsafe_allow_html=True)
        
        conn = get_conn()
        t_files = pd.read_sql_query("SELECT COUNT(*) c FROM files", conn).iloc[0]['c']
        t_analyses = pd.read_sql_query("SELECT COUNT(*) c FROM analyses", conn).iloc[0]['c']
        t_findings = pd.read_sql_query("SELECT COUNT(*) c FROM findings", conn).iloc[0]['c']
        conn.close()

        c1, c2, c3 = st.columns(3)
        c1.metric("📁 File Diproses", t_files)
        c2.metric("📊 Analisis Selesai", t_analyses)
        c3.metric("⚠️ Total Temuan Risiko", t_findings)

        st.markdown("---")
        st.subheader("💡 Panduan Penggunaan Sistem")
        st.markdown("""
        1. **Fitur 1 (Audit Profesional):** Unggah file transaksi Excel/CSV Anda untuk mendeteksi duplikasi, nilai negatif, transaksi akhir pekan, hingga anomali *outlier*.
        2. **Fitur 2 (Analisis Keuangan):** Evaluasi rasio profitabilitas, likuiditas, solvabilitas, atau rasio fiskal daerah secara otomatis.
        3. **Fitur 3 (Cash Flow Intelligence):** Pantau saldo berjalan (*running balance*), ketahanan kas (*cash runway*), hingga simulasi AI *What-If*.
        """)

    # ---------------- 🔎 FITUR 1: AUDIT ----------------
    elif menu == "🔎 Fitur 1: Audit Profesional":
        st.markdown("""
        <div class="finora-header">
            <h1>🔎 Fitur 1: Audit Profesional & Deteksi Anomali</h1>
            <p>Unggah file transaksi dan jalankan uji kepatuhan otomatis.</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded = st.file_uploader("📂 Unggah File Transaksi (Excel / CSV)", type=["xlsx", "csv"], key="audit_upload")
        if uploaded:
            df = pd.read_excel(uploaded) if uploaded.name.endswith('xlsx') else pd.read_csv(uploaded)
            detected = auto_detect_columns(df)
            
            with st.expander("⚙️ Pengaturan & Pemetaan Kolom"):
                c1, c2, c3 = st.columns(3)
                date_col = c1.selectbox("Kolom Tanggal", [""] + list(df.columns), index=list(df.columns).index(detected["date"])+1 if detected["date"] else 0)
                amount_col = c2.selectbox("Kolom Nominal", [""] + list(df.columns), index=list(df.columns).index(detected["amount"])+1 if detected["amount"] else 0)
                desc_col = c3.selectbox("Kolom Deskripsi", [""] + list(df.columns), index=list(df.columns).index(detected["desc"])+1 if detected["desc"] else 0)

            entitas = st.text_input("Nama Entitas / Perusahaan", "PT Maju Sejahtera")
            
            if st.button("🚀 Jalankan Prosedur Audit", use_container_width=True):
                all_findings, flagged_data = [], {}
                cols = [c for c in [date_col, amount_col, desc_col] if c]
                f, flag = audit_empty(df, cols); all_findings.extend(f); flagged_data["Kelengkapan Kosong"] = flag
                if amount_col:
                    f, flag = audit_negative(df, amount_col); all_findings.extend(f); flagged_data["Nilai Negatif"] = flag
                    f, flag = audit_outliers_iqr(df, amount_col); all_findings.extend(f); flagged_data["Outlier IQR"] = flag
                
                score = min(100, len(all_findings) * 15)
                st.success(f"Audit Selesai! Ditemukan {len(all_findings)} potensi temuan.")
                
                st.markdown("### Ringkasan Temuan Audit")
                for item in all_findings:
                    st.warning(f"**[{item['risk_level']}] {item['module']}**: {item['title']} — *Rekomendasi:* {item['recommendation']}")

    # ---------------- 📈 FITUR 2: ANALISIS KEUANGAN ----------------
    elif menu == "📈 Fitur 2: Analisis Laporan Keuangan":
        st.markdown("""
        <div class="finora-header">
            <h1>📈 Fitur 2: Analisis Laporan Keuangan & Rasio</h1>
            <p>Ekstraksi otomatis laporan posisi keuangan dan perhitungan rasio penting.</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_fin = st.file_uploader("📂 Unggah Laporan Keuangan (Excel / CSV)", type=["xlsx", "csv"], key="fin_upload")
        if uploaded_fin:
            df_fin = pd.read_excel(uploaded_fin) if uploaded_fin.name.endswith('xlsx') else pd.read_csv(uploaded_fin)
            if st.button("📊 Hitung Rasio & Analisis", use_container_width=True):
                nilai_utama = auto_extract_financials(df_fin, entity_type)
                rasio = hitung_rasio_dari_nilai(nilai_utama, entity_type)
                
                st.success("Analisis Keuangan Berhasil Dijalankan!")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("### Komponen Utama")
                    for k, v in nilai_utama.items():
                        if v > 0: st.metric(k.replace("_", " ").title(), format_rupiah(v))
                with c2:
                    st.markdown("### Rasio Keuangan")
                    for k, v in rasio.items():
                        st.metric(k, f"{v}")

    # ---------------- 💵 FITUR 3: CASH FLOW INTELLIGENCE ----------------
    elif menu == "💵 Fitur 3: Cash Flow Intelligence":
        st.markdown("""
        <div class="finora-header">
            <h1>💵 Fitur 3: Cash Flow Intelligence (Arus Kas)</h1>
            <p>Analisis likuiditas, saldo berjalan, burn rate, runway, dan Konsultasi AI.</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_cf = st.file_uploader("📂 Unggah Data Mutasi Kas / Buku Besar (Excel / CSV)", type=["xlsx", "csv"], key="cf_upload")
        if uploaded_cf:
            df_cf = pd.read_excel(uploaded_cf) if uploaded_cf.name.endswith('xlsx') else pd.read_csv(uploaded_cf)
            detected_cf = auto_detect_columns(df_cf)
            
            with st.expander("⚙️ Pemetaan Kolom Arus Kas"):
                c1, c2, c3, c4 = st.columns(4)
                cf_date = c1.selectbox("Kolom Tanggal", [""] + list(df_cf.columns), index=list(df_cf.columns).index(detected_cf["date"])+1 if detected_cf["date"] else 0)
                cf_desc = c2.selectbox("Kolom Keterangan", [""] + list(df_cf.columns), index=list(df_cf.columns).index(detected_cf["desc"])+1 if detected_cf["desc"] else 0)
                cf_in = c3.selectbox("Kolom Kas Masuk", [""] + list(df_cf.columns), index=list(df_cf.columns).index(detected_cf["cash_in"] or detected_cf["amount"])+1 if (detected_cf["cash_in"] or detected_cf["amount"]) else 0)
                cf_out = c4.selectbox("Kolom Kas Keluar", [""] + list(df_cf.columns), index=list(df_cf.columns).index(detected_cf["cash_out"])+1 if detected_cf["cash_out"] else 0)
            
            initial_balance = st.number_input("Saldo Awal Kas (Rp)", value=10_000_000.0, step=1_000_000.0)
            
            if st.button("🚀 Analisis Arus Kas & Likuiditas", use_container_width=True):
                cf_df, cf_metrics, cf_anomalies = process_cash_flow_data(df_cf, cf_date, cf_desc, cf_in, cf_out, initial_balance)
                
                st.session_state['cf_results'] = {
                    "df": cf_df, "metrics": cf_metrics, "anomalies": cf_anomalies, "filename": uploaded_cf.name
                }
                st.success("Analisis Arus Kas Berhasil!")

        if 'cf_results' in st.session_state:
            res = st.session_state['cf_results']
            m = res["metrics"]
            
            st.markdown("---")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Saldo Kas Akhir", format_rupiah(m["saldo_akhir"]))
            c2.metric("Total Kas Masuk", format_rupiah(m["total_masuk"]))
            c3.metric("Total Kas Keluar", format_rupiah(m["total_keluar"]))
            c4.metric("Cash Runway", f"{m['cash_runway_months']} Bulan")

            tab_cf_grafik, tab_cf_ai = st.tabs(["📈 Grafik Tren & Anomali", "🤖 Tanya AI & Simulasi What-If"])
            
            with tab_cf_grafik:
                if not res["df"].empty and 'Running Balance' in res["df"].columns:
                    fig = px.line(res["df"], x=res["df"].columns[0], y="Running Balance", title="Tren Saldo Berjalan (Running Balance)", markers=True)
                    st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("### Deteksi Anomali Pengeluaran Kas")
                if not res["anomalies"].empty:
                    st.warning(f"Ditemukan {len(res['anomalies']} transaksi pengeluaran di atas batas normal.")
                    st.dataframe(res["anomalies"], use_container_width=True)
                else:
                    st.info("Tidak ada anomali pengeluaran ekstrem.")

            with tab_cf_ai:
                st.markdown("### 🤖 AI Cash Flow Advisor & Simulasi What-If")
                if "cf_chat_history" not in st.session_state: st.session_state.cf_chat_history = []
                
                for msg in st.session_state.cf_chat_history:
                    with st.chat_message(msg["role"]): st.markdown(msg["content"])
                
                prompt_cf = st.text_input("Tanyakan kondisi kas atau simulasi (misal: 'Jika saya memotong biaya operasional 20%, berapa lama runway bertambah?')", key="cf_prompt_input")
                if st.button("Kirim ke AI", key="btn_send_cf_ai"):
                    if prompt_cf:
                        st.session_state.cf_chat_history.append({"role": "user", "content": prompt_cf})
                        context_cf_ai = f"Metrik Kas: Saldo Akhir Rp {m['saldo_akhir']:,.0f}, Burn Rate Rp {m['monthly_burn_rate']:,.0f}/bln, Runway {m['cash_runway_months']} bulan."
                        with st.chat_message("assistant"):
                            response_ai = st.write_stream(ask_gemini_stream(prompt_cf, context_cf_ai, system_role="AI Cash Flow Advisor & Senior Financial Expert"))
                        st.session_state.cf_chat_history.append({"role": "assistant", "content": response_ai})
                        st.rerun()

    # ---------------- 💾 ARSIP ----------------
    elif menu == "💾 Arsip & Riwayat Laporan":
        st.markdown("""
        <div class="finora-header">
            <h1>💾 Arsip & Riwayat Laporan</h1>
            <p>Daftar seluruh analisis yang pernah disimpan ke database lokal.</p>
        </div>
        """, unsafe_allow_html=True)
        df_hist = fetch_analyses()
        if df_hist.empty:
            st.info("Belum ada riwayat laporan tersimpan.")
        else:
            st.dataframe(df_hist, use_container_width=True)
