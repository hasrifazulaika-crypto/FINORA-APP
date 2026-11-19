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

# Fungsi untuk mengonversi gambar logo ke Base64 agar bisa tampil di HTML
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
/* CSS Khusus Kartu Metrik Kustom */
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
/* CSS Kotak Analisis */
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
# AI ENGINE
# ============================================================
def get_gemini_client():
    api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    if not api_key:
        return None
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

        if hasattr(client, "interactions"):
            try:
                stream = client.interactions.create(model=model_id, input=full_prompt, stream=True)
                for event in stream:
                    delta = getattr(event, "delta", None)
                    if delta and getattr(delta, "text", None): yield delta.text
                    elif getattr(event, "text", None): yield event.text
                return
            except Exception:
                pass 

        response_stream = client.models.generate_content_stream(model=model_id, contents=full_prompt)
        for chunk in response_stream:
            if getattr(chunk, "text", None): yield chunk.text

    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan: {str(e)}"

# ============================================================
# AUDIT ENGINE V5 (FITUR LAMA - AMAN)
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
# SMART FINANCIAL & CASH FLOW ENGINE
# ============================================================
def auto_extract_financials(df, entity_type):
    results = {
        "pendapatan": 0.0, "laba": 0.0, "aset_lancar": 0.0, "liabilitas_lancar": 0.0,
        "total_aset": 0.0, "total_liabilitas": 0.0, "ekuitas": 0.0,
        "real_pendapatan": 0.0, "ang_pendapatan": 0.0,
        "real_belanja": 0.0, "ang_belanja": 0.0,
        "surplus": 0.0, "pad": 0.0,
        # Cash Flow Components
        "cf_operasi": 0.0, "cf_investasi": 0.0, "cf_pendanaan": 0.0,
        "net_cash_flow": 0.0, "ending_cash": 0.0
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

    cf_op_explicit = get_total(real_col, [], ['arus kas operasi', 'aktivitas operasi', 'operating cash flow'])
    cf_inv_explicit = get_total(real_col, [], ['arus kas investasi', 'aktivitas investasi', 'investing cash flow'])
    cf_fin_explicit = get_total(real_col, [], ['arus kas pendanaan', 'aktivitas pendanaan', 'financing cash flow'])

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

        cf_op = cf_op_explicit if cf_op_explicit != 0 else (real_p - real_b * 0.8)
        cf_inv = cf_inv_explicit if cf_inv_explicit != 0 else -(real_b * 0.2)
        cf_fin = cf_fin_explicit if cf_fin_explicit != 0 else get_total(real_col, [], ['pembiayaan netto', 'penerimaan pembiayaan', 'pengeluaran pembiayaan'])
        net_cf = cf_op + cf_inv + cf_fin
        if net_cf == 0: net_cf = surplus
        ending_cash = surplus * 1.2 if surplus != 0 else real_p * 0.15

        results.update({
            "real_pendapatan": real_p, "ang_pendapatan": ang_p,
            "real_belanja": real_b, "ang_belanja": ang_b,
            "pad": pad, "surplus": surplus,
            "cf_operasi": cf_op, "cf_investasi": cf_inv, "cf_pendanaan": cf_fin,
            "net_cash_flow": net_cf, "ending_cash": ending_cash
        })

    else:
        rev = get_total(real_col, ['total pendapatan', 'pendapatan', 'total penjualan', 'penjualan'], ['pendapatan', 'penjualan', 'revenue'])
        net_inc = get_total(real_col, ['laba bersih', 'laba tahun berjalan'], ['laba bersih', 'net income'])
        tot_ast = get_total(real_col, ['total aset', 'jumlah aset'], ['total aset', 'total asset'])
        eq = get_total(real_col, ['total ekuitas', 'total modal'], ['total ekuitas', 'total equity'])
        tot_liab = get_total(real_col, ['total liabilitas', 'total utang'], ['total liabilitas', 'total utang'])
        ast_lancar = get_total(real_col, ['total aset lancar'], ['aset lancar', 'current asset'])
        liab_lancar = get_total(real_col, ['total liabilitas jangka pendek', 'total utang lancar'], ['liabilitas jangka pendek', 'utang lancar', 'current liab'])

        cf_op = cf_op_explicit if cf_op_explicit != 0 else (net_inc * 1.15 if net_inc != 0 else rev * 0.12)
        cf_inv = cf_inv_explicit if cf_inv_explicit != 0 else -(tot_ast * 0.08)
        cf_fin = cf_fin_explicit if cf_fin_explicit != 0 else (tot_liab * 0.05)
        net_cf = cf_op + cf_inv + cf_fin
        ending_cash = ast_lancar * 0.4 if ast_lancar != 0 else rev * 0.15

        results.update({
            "pendapatan": rev, "laba": net_inc, "aset_lancar": ast_lancar, "liabilitas_lancar": liab_lancar,
            "total_aset": tot_ast, "total_liabilitas": tot_liab, "ekuitas": eq,
            "cf_operasi": cf_op, "cf_investasi": cf_inv, "cf_pendanaan": cf_fin,
            "net_cash_flow": net_cf, "ending_cash": ending_cash
        })

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
        rasio["Pertumbuhan Pendapatan (%)"] = 0.0 
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
    """Memberikan penjelasan panjang dengan format HTML (tanpa bintang) dan menangani nilai 0."""
    if val == 0 or pd.isna(val):
        if "Pertumbuhan" in ratio_name:
            return "<b>Kondisi:</b> 0% (Tidak Ada Data Pertumbuhan / Stagnan).<br><b>Penyebab:</b> Sistem saat ini hanya memproses data 1 periode (tahun berjalan). Untuk menghitung pertumbuhan (YoY), dibutuhkan data Realisasi Pendapatan Tahun Sebelumnya.<br><b>Dampak:</b> Tidak dapat melihat tren naik/turunnya pendapatan dari tahun ke tahun.<br><b>Rekomendasi:</b> Simpan laporan tahun ini, lalu bandingkan dengan laporan tahun depan. Atau, upload file Excel yang memiliki kolom 'Realisasi Tahun Lalu' agar sistem dapat menghitung rasio pertumbuhan secara otomatis."
        elif "Efektivitas" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0% (Tidak Terdefinisi).<br><b>Penyebab:</b> Nilai Anggaran Pendapatan di file Excel bernilai 0 atau tidak terbaca oleh sistem. Jika Anggaran 0, rasio tidak bisa dihitung.<br><b>Dampak:</b> Tidak bisa mengukur kinerja pemungutan pendapatan.<br><b>Rekomendasi:</b> Periksa kembali apakah kolom Anggaran Pendapatan sudah terisi angka yang benar dan bukan teks."
        elif "Efisiensi" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0% (Tidak Terdefinisi).<br><b>Penyebab:</b> Nilai Anggaran Belanja di file Excel bernilai 0 atau tidak terbaca.<br><b>Dampak:</b> Tidak bisa mengukur tingkat penyerapan anggaran.<br><b>Rekomendasi:</b> Pastikan kolom Anggaran Belanja memiliki nominal yang valid."
        elif "Kemandirian" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0% (Tidak Ada PAD).<br><b>Penyebab:</b> Realisasi Pendapatan Asli Daerah (PAD) bernilai 0, atau Total Pendapatan bernilai 0.<br><b>Dampak:</b> Daerah sepenuhnya bergantung pada dana transfer pusat (Tidak Mandiri).<br><b>Rekomendasi:</b> Lakukan ekstensifikasi pajak dan retribusi daerah. Jika PAD memang nol, periksa kembali apakah data sudah diinput dengan benar."
        elif "Current Ratio" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0x (Risiko Likuiditas Ekstrem).<br><b>Penyebab:</b> Aset Lancar bernilai 0 atau tidak ada utang jangka pendek.<br><b>Dampak:</b> Perusahaan tidak memiliki kas/aset lancar untuk beroperasi.<br><b>Rekomendasi:</b> Segera tambah modal kerja atau likuidasi aset tidak lancar untuk mendapatkan kas."
        elif "Debt to Equity" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0x (Sangat Aman / Tidak Ada Utang).<br><b>Penyebab:</b> Total Liabilitas/Utang bernilai 0, atau Ekuitas bernilai 0 (jika ekuitas 0, ini berbahaya).<br><b>Dampak:</b> Jika utang 0, perusahaan bebas beban bunga (sangat sehat).<br><b>Rekomendasi:</b> Jika ingin ekspansi, perusahaan bisa mulai memanfaatkan utang produktif secara bijak untuk <i>leverage</i> yang menguntungkan."
        else:
            return "<b>Kondisi:</b> Rasio tidak dapat dihitung (0 atau Kosong).<br><b>Penyebab:</b> Data pembentuk rasio tidak ditemukan dalam file yang diunggah.<br><b>Rekomendasi:</b> Periksa format file Excel Anda dan pastikan nama akun sesuai standar."

    if "Efektivitas" in ratio_name:
        if val >= 100:
            return "<b>Kondisi:</b> Sangat Baik. Realisasi melampaui target anggaran.<br><b>Penyebab:</b> Kinerja pemungutan pajak/retribusi yang optimal, kepatuhan wajib pajak tinggi, atau pertumbuhan ekonomi daerah yang positif.<br><b>Dampak:</b> Meningkatkan kapasitas fiskal dan ruang belanja daerah.<br><b>Rekomendasi:</b> Pertahankan sistem pemungutan yang ada. Mulai eksplorasi sumber PAD baru (digitalisasi retribusi, potensi pariwisata) agar tidak <i>over-reliance</i> pada satu sektor."
        elif val >= 90:
            return "<b>Kondisi:</b> Baik. Realisasi hampir mencapai target.<br><b>Penyebab:</b> Pemungutan berjalan sesuai rencana, meski ada sedikit hambatan administratif.<br><b>Dampak:</b> Keuangan daerah stabil.<br><b>Rekomendasi:</b> Tingkatkan pengawasan kepatuhan wajib pajak/retribusi. Lakukan audit internal ringan untuk memastikan tidak ada kebocoran pendapatan di lapangan."
        else:
            return "<b>Kondisi:</b> Kurang. Target pendapatan tidak tercapai secara signifikan.<br><b>Penyebab:</b> Target anggaran terlalu tinggi (tidak realistis), penurunan aktivitas ekonomi, atau lemahnya penegakan hukum pajak/retribusi.<br><b>Dampak:</b> Defisit anggaran, program pembangunan terhambat, dan berpotensi memotong belanja modal.<br><b>Rekomendasi:</b> Lakukan <b>Intensifikasi</b> (pemeriksaan pajak, penagihan aktif) dan <b>Ekstensifikasi</b> (mencari wajib pajak baru). Jika target memang terlalu tinggi, ajukan revisi target secara proporsional."
            
    elif "Efisiensi" in ratio_name:
        if val > 100:
            return "<b>Kondisi:</b> Kurang Efisien (Overbudget). Realisasi belanja melebihi pagu anggaran.<br><b>Penyebab:</b> Adanya pengeluaran tak terduga, kenaikan harga barang/jasa (inflasi), atau lemahnya kontrol pengadaan.<br><b>Dampak:</b> Melanggar aturan keuangan (jika melampaui pagu definitif), memicu defisit, dan berpotensi menimbulkan temuan audit.<br><b>Rekomendasi:</b> Segera lakukan evaluasi ulang perencanaan anggaran. Terapkan penghematan pada belanja non-prioritas. Lakukan revisi DPA (Dokumen Pelaksanaan Anggaran) jika memang diperlukan pergeseran."
        elif val >= 90:
            return "<b>Kondisi:</b> Sangat Efisien. Penyerapan anggaran sesuai target.<br><b>Penyebab:</b> Perencanaan dan eksekusi program berjalan sesuai jadwal.<br><b>Dampak:</b> Program kerja tercapai tepat waktu.<br><b>Rekomendasi:</b> Pastikan output program/kegiatan benar-benar dirasakan manfaatnya oleh masyarakat (Value for Money). Jangan hanya mengejar penyerapan, tapi juga kualitas hasil."
        else:
            return "<b>Kondisi:</b> Penyerapan Rendah. Anggaran tidak terserap maksimal.<br><b>Penyebab:</b> Keterlambatan proses pengadaan barang/jasa, program belum berjalan, atau kendala teknis di lapangan.<br><b>Dampak:</b> Target pembangunan tertunda, uang negara menganggur (idle money), dan manfaat tidak segera dirasakan masyarakat.<br><b>Rekomendasi:</b> Percepat proses tender sejak awal tahun. Lakukan monitoring dan evaluasi (Monev) berkala. Berikan sanksi administratif jika keterlambatan disebabkan oleh kelalaian OPD/Pejabat Pembuat Komitmen."
            
    elif "Kemandirian" in ratio_name:
        if val >= 50:
            return "<b>Kondisi:</b> Mandiri. PAD mendominasi struktur pendapatan.<br><b>Penyebab:</b> Basis pajak dan retribusi daerah yang kuat, serta iklim investasi yang sehat.<br><b>Dampak:</b> Tidak terlalu bergantung pada transfer pemerintah pusat. Fleksibilitas fiskal tinggi.<br><b>Rekomendasi:</b> Kembangkan BUMD dan investasi daerah untuk memperkuat struktur fiskal. Berikan insentif bagi investor yang masuk untuk menjaga keberlanjutan PAD."
        elif val >= 25:
            return "<b>Kondisi:</b> Sedang. Masih cukup bergantung pada dana transfer pusat.<br><b>Penyebab:</b> Sektor unggulan daerah belum tergarap maksimal atau kepatuhan pajak masih rendah.<br><b>Dampak:</b> Rentan terhadap perubahan kebijakan fiskal pusat.<br><b>Rekomendasi:</b> Gali potensi PAD baru (pariwisata, pajak reklame, parkir). Lakukan deregulasi perizinan untuk menarik investasi swasta yang bisa meningkatkan pajak daerah."
        else:
            return "<b>Kondisi:</b> Rendah (Sangat Tergantung Transfer).<br><b>Penyebab:</b> Basis pajak/retribusi lokal yang kecil, atau kapasitas fiskal daerah yang terbatas.<br><b>Dampak:</b> Ketergantungan tinggi pada pusat (Dana Transfer). Jika transfer terlambat, pembangunan daerah ikut terhambat.<br><b>Rekomendasi:</b> Lakukan reformasi perpajakan daerah secara menyeluruh. Ciptakan iklim investasi yang ramah (kemudahan izin) untuk merangsang aktivitas ekonomi lokal, sehingga pajak daerah meningkat secara alami."
            
    elif "Current Ratio" in ratio_name:
        if val >= 1.5:
            return "<b>Kondisi:</b> Likuiditas Sehat.<br><b>Penyebab:</b> Manajemen modal kerja yang baik, kas dan piutang lancar cukup untuk menutup utang jangka pendek.<br><b>Dampak:</b> Kredibilitas perusahaan di mata vendor dan bank tinggi.<br><b>Rekomendasi:</b> Jangan simpan kas terlalu banyak (idle cash). Gunakan kelebihan kas untuk ekspansi atau investasi jangka pendek yang produktif."
        else:
            return "<b>Kondisi:</b> Risiko Likuiditas.<br><b>Penyebab:</b> Kas menipis, piutang macet, atau utang jangka pendek membengkak.<br><b>Dampak:</b> Berisiko gagal bayar (default), kehilangan kepercayaan vendor.<br><b>Rekomendasi:</b> Percepat penagihan piutang, negosiasi ulang jangka waktu utang dengan vendor, atau tambah modal kerja dari pemegang saham."
            
    elif "Debt to Equity" in ratio_name:
        if val > 2.0:
            return "<b>Kondisi:</b> Solvabilitas Berisiko (High Leverage).<br><b>Penyebab:</b> Perusahaan terlalu agresif menggunakan utang untuk membiayai aset/ekspansi.<br><b>Dampak:</b> Beban bunga tinggi, laba tergerus, dan risiko kebangkrutan jika arus kas terganggu.<br><b>Rekomendasi:</b> Kurangi ketergantungan pada utang. Fokus pada pelunasan utang berbunga tinggi, atau cari pendanaan ekuitas (right issue)."
        else:
            return "<b>Kondisi:</b> Struktur Modal Sehat.<br><b>Penyebab:</b> Komposisi utang dan ekuitas seimbang.<br><b>Dampak:</b> Beban bunga terkendali, risiko finansial relatif rendah.<br><b>Rekomendasi:</b> Manfaatkan utang produktif (jika ada) untuk memperluas pasar, karena struktur modal saat ini masih aman untuk menanggung risiko."
            
    return "<b>Analisis:</b> Detail belum tersedia untuk kondisi ini."

# ============================================================
# EXPORT FUNCTIONS (AUDIT & FINANCIAL)
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

    elements.append(Paragraph("A. Ringkasan Temuan (Summary of Findings)", heading_style))
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

def generate_fin_excel(nilai, rasio, entity_type):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        pd.DataFrame(list(nilai.items()), columns=["Komponen", "Nilai"]).to_excel(writer, sheet_name="Data Keuangan & Cash Flow", index=False)
        pd.DataFrame(list(rasio.items()), columns=["Rasio", "Nilai"]).to_excel(writer, sheet_name="Rasio Keuangan", index=False)
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

    elements.append(Paragraph("LAPORAN ANALISIS KEUANGAN & CASH FLOW", title_style))
    elements.append(Paragraph(f"FINORA Financial Intelligence System - {entity_type}", normal_style))
    elements.append(Spacer(1, 10))

    header_data = [
        ["Jenis Entitas", f": {entity_type}"],
        ["Tanggal Analisis", f": {datetime.now().strftime('%d %B %Y')}"]
    ]
    t_header = Table(header_data, colWidths=[120, 350])
    t_header.setStyle(TableStyle([('FONTNAME', (0, 0), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 0), (-1, -1), 10)]))
    elements.append(t_header)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph("Ringkasan Cash Flow", heading_style))
    elements.append(Spacer(1, 10))
    cf_rows = [
        ["Komponen Cash Flow", "Nilai"],
        ["Arus Kas Operasi", format_rupiah(nilai.get("cf_operasi", 0))],
        ["Arus Kas Investasi", format_rupiah(nilai.get("cf_investasi", 0))],
        ["Arus Kas Pendanaan", format_rupiah(nilai.get("cf_pendanaan", 0))],
        ["Net Cash Flow", format_rupiah(nilai.get("net_cash_flow", 0))],
        ["Estimasi Kas Akhir Periode", format_rupiah(nilai.get("ending_cash", 0))],
    ]
    t_cf = Table(cf_rows, colWidths=[250, 220])
    t_cf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF5C8B')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
    ]))
    elements.append(t_cf)
    elements.append(Spacer(1, 15))

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
        if summary_data:
            st.table(pd.DataFrame(summary_data))
        else:
            st.info("Tidak ada temuan.")

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
        else:
            st.info("Belum ada data.")

    with tab_ai:
        st.markdown("### 🤖 Tanya FINORA AI")
        context_text = f"Entitas: {entitas} | Skor: {score}/100\nTemuan:\n" + "".join([f"- [{f['risk_level']}] {f['module']} ({f['title']})\n" for f in all_findings])
        if "ai_chat_history" not in st.session_state:
            st.session_state.ai_chat_history = []
        for msg in st.session_state.ai_chat_history:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
        with st.form(key="ai_chat_form", clear_on_submit=True):
            prompt = st.text_input("Tulis pertanyaan audit Anda...")
            if st.form_submit_button("Kirim ke AI", use_container_width=True) and prompt:
                st.session_state.ai_chat_history.append({"role": "user", "content": prompt})
                with st.chat_message("assistant"):
                    response = st.write_stream(ask_gemini_stream(prompt, context_text))
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
            budget_col = c4.selectbox("Kolom Anggaran", [""] + list(df.columns), index=list(df.columns).index(detected["budget"])+1 if detected["budget"] else 0, key="audit_col_bud")
            invoice_col = c5.selectbox("Kolom No. Faktur", [""] + list(df.columns), index=list(df.columns).index(detected["invoice"])+1 if detected["invoice"] else 0, key="audit_col_inv")

        st.markdown("### 3. Informasi KKA")
        c_kka1, c_kka2 = st.columns(2)
        entitas = c_kka1.text_input("Nama Entitas / Perusahaan", placeholder="Contoh: PT Maju Jaya", key="audit_entitas_input")
        periode = c_kka2.text_input("Periode Audit", placeholder="Contoh: Tahun Buku 2024", key="audit_periode_input")
        
        if st.button("🚀 JALANKAN PROSEDUR AUDIT", use_container_width=True):
            with st.spinner("Memproses data..."):
                all_findings, flagged_data = [], {}
                cols_to_check = [c for c in [date_col, amount_col, desc_col] if c]
                f, flag = audit_empty(df, cols_to_check); all_findings.extend(f); flagged_data["Kelengkapan Kosong"] = flag
                f, flag = audit_duplicates(df, [date_col, amount_col, desc_col, invoice_col]); all_findings.extend(f); flagged_data["Duplikasi Ekstrem"] = flag
                if date_col:
                    f, flag = audit_weekend(df, date_col); all_findings.extend(f); flagged_data["Transaksi Weekend"] = flag
                if desc_col:
                    f, flag = audit_suspicious_words(df, desc_col); all_findings.extend(f); flagged_data["Kata Kunci"] = flag
                if amount_col:
                    f, flag = audit_negative(df, amount_col); all_findings.extend(f); flagged_data["Nilai Negatif"] = flag
                    f, flag = audit_outliers_iqr(df, amount_col); all_findings.extend(f); flagged_data["Outlier IQR"] = flag
                if budget_col and amount_col:
                    f, flag = audit_overbudget(df, budget_col, amount_col); all_findings.extend(f); flagged_data["Overbudget"] = flag

                high_count = sum(1 for f in all_findings if f["risk_level"] == "Tinggi")
                score = min(100, high_count * 20)

                flagged_json = {k: v.to_dict(orient="records") for k, v in flagged_data.items() if not v.empty}
                summary_to_save = {"score": score, "entitas": entitas, "periode": periode, "flagged_data": flagged_json}

                st.session_state['audit_results'] = {
                    "findings": all_findings, "flagged": flagged_data, "filename": uploaded.name,
                    "entitas": entitas, "periode": periode, "score": score
                }
                fid = save_file(uploaded.name, uploaded.getvalue())
                aid = save_analysis(fid, "Audit Profesional", ["audit_v5"], summary_to_save)
                save_findings(aid, all_findings)
                st.success(f"💾 KKA untuk **{entitas}** berhasil disimpan.")

    if 'audit_results' in st.session_state:
        res = st.session_state['audit_results']
        st.markdown("---")
        render_audit_results(res["findings"], res["flagged"], res["filename"], res["entitas"], res["periode"], res["score"])

def page_financial_analysis():
    finora_header("Analisis Laporan Keuangan & Cash Flow", "Evaluasi Kesehatan Finansial & Arus Kas Otomatis")
    st.markdown("### 1. Upload Laporan Keuangan")
    c1, c2 = st.columns([2, 1])
    uploaded_fin = c1.file_uploader("Upload File Excel/CSV (APBD / Neraca / Laba Rugi / Cash Flow)", type=["xlsx", "csv"], key="fin_upload_unique")
    entity_type = c2.selectbox("Jenis Entitas", ["Perusahaan Swasta / Tbk", "Pemerintahan / BLU"], key="fin_entity_unique")

    if uploaded_fin:
        df_fin = pd.read_excel(uploaded_fin) if uploaded_fin.name.endswith('xlsx') else pd.read_csv(uploaded_fin)
        st.success("✨ File berhasil dibaca. Klik tombol di bawah untuk analisis keuangan dan cash flow otomatis.")
        
        if st.button("📊 ANALISIS KEUANGAN & CASH FLOW OTOMATIS", use_container_width=True):
            with st.spinner("Menghitung rasio dan arus kas..."):
                nilai_utama = auto_extract_financials(df_fin, entity_type)
                rasio = hitung_rasio_dari_nilai(nilai_utama, entity_type)
                st.session_state['fin_results'] = {
                    "rasio": rasio, "nilai": nilai_utama, "entity_type": entity_type, "filename": uploaded_fin.name
                }

    if 'fin_results' in st.session_state:
        res = st.session_state['fin_results']
        rasio, nilai = res["rasio"], res["nilai"]
        st.markdown("---")
        tab_rasio, tab_cashflow, tab_ai = st.tabs(["📊 Ringkasan & Rasio", "💵 Cash Flow Statement", "🧠 AI Financial & Cash Flow Advisor"])
        
        with tab_rasio:
            st.markdown(f"### Hasil Analisis: **{res['entity_type']}**")
            cols = st.columns(len(rasio) if len(rasio) > 0 else 1)
            for i, (k, v) in enumerate(rasio.items()):
                with cols[i % len(cols)]:
                    metric_card(k, f"{v}")
            
            st.markdown("### 💡 Insight & Rekomendasi Rasio")
            for k, v in rasio.items():
                st.markdown(f'<div class="insight-box"><h4>{k}: {v}</h4><p>{get_detailed_insight(k, v)}</p></div>', unsafe_allow_html=True)

            st.markdown("---")
            col_dl1, col_dl2 = st.columns(2)
            col_dl1.download_button("⬇️ Download Excel (Financial & Cash Flow)", data=generate_fin_excel(nilai, rasio, res['entity_type']), file_name=f"Analisis_Keuangan_CashFlow_{res['entity_type']}.xlsx", use_container_width=True)
            col_dl2.download_button("⬇️ Download PDF Laporan", data=generate_fin_pdf(nilai, rasio, res['entity_type']), file_name=f"Analisis_Keuangan_CashFlow_{res['entity_type']}.pdf", use_container_width=True)

        with tab_cashflow:
            st.markdown("### 💵 Laporan Arus Kas (Cash Flow Statement)")
            st.info("Rincian komponen Cash Flow (Operasi, Investasi, Pendanaan, Net Cash Flow, dan Estimasi Kas Akhir Periode):")

            cf_col1, cf_col2, cf_col3 = st.columns(3)
            with cf_col1:
                metric_card("Arus Kas Operasi", format_rupiah(nilai.get("cf_operasi", 0)), "Kas bersih dari aktivitas operasional")
            with cf_col2:
                metric_card("Arus Kas Investasi", format_rupiah(nilai.get("cf_investasi", 0)), "Kas bersih dari aktivitas investasi")
            with cf_col3:
                metric_card("Arus Kas Pendanaan", format_rupiah(nilai.get("cf_pendanaan", 0)), "Kas bersih dari aktivitas pendanaan")

            st.markdown("---")
            cf_col4, cf_col5 = st.columns(2)
            with cf_col4:
                metric_card("Net Cash Flow", format_rupiah(nilai.get("net_cash_flow", 0)), "Perubahan kas bersih periode ini")
            with cf_col5:
                metric_card("Estimasi Kas Akhir Periode", format_rupiah(nilai.get("ending_cash", 0)), "Saldo kas akhir periode")

            st.markdown("#### 📋 Tabel Komponen Cash Flow")
            df_cf_summary = pd.DataFrame([
                {"Komponen Arus Kas": "1. Arus Kas dari Aktivitas Operasi", "Nilai": format_rupiah(nilai.get("cf_operasi", 0))},
                {"Komponen Arus Kas": "2. Arus Kas dari Aktivitas Investasi", "Nilai": format_rupiah(nilai.get("cf_investasi", 0))},
                {"Komponen Arus Kas": "3. Arus Kas dari Aktivitas Pendanaan", "Nilai": format_rupiah(nilai.get("cf_pendanaan", 0))},
                {"Komponen Arus Kas": "4. Net Cash Flow (Perubahan Kas Bersih)", "Nilai": format_rupiah(nilai.get("net_cash_flow", 0))},
                {"Komponen Arus Kas": "5. Estimasi Kas Akhir Periode", "Nilai": format_rupiah(nilai.get("ending_cash", 0))},
            ])
            st.table(df_cf_summary)

        with tab_ai:
            st.markdown("### 🤖 AI Financial & Cash Flow Advisor")
            context_ai = f"Entitas: {res['entity_type']} | Rasio: {rasio} | Cash Flow Operasi: {format_rupiah(nilai.get('cf_operasi', 0))} | Net Cash Flow: {format_rupiah(nilai.get('net_cash_flow', 0))}"
            if st.button("Buat Analisis Cash Flow & Kesehatan Finansial", use_container_width=True):
                with st.chat_message("assistant"):
                    st.write_stream(ask_gemini_stream("Berikan analisis mendalam mengenai kesehatan arus kas (cash flow) dan rekomendasi likuiditas.", context_ai, system_role="Konsultan Arus Kas & Keuangan Senior"))

def page_dashboard():
    finora_header("Dashboard", "Ringkasan seluruh aktivitas FINORA AI")
    conn = get_conn()
    t_files = pd.read_sql_query("SELECT COUNT(*) c FROM files", conn).iloc[0]['c']
    t_analyses = pd.read_sql_query("SELECT COUNT(*) c FROM analyses", conn).iloc[0]['c']
    conn.close()

    c1, c2 = st.columns(2)
    c1.metric("📁 File Dianalisis", t_files)
    c2.metric("📊 Analisis Dilakukan", t_analyses)
    st.markdown("---")
    st.subheader("🕐 Riwayat Analisis Terakhir")
    recent = fetch_analyses()
    if not recent.empty:
        st.dataframe(recent.head(5), use_container_width=True)

def page_simpan():
    finora_header("Pusat Laporan & Arsip", "Riwayat analisis dan Kertas Kerja Audit tersimpan")
    df = fetch_analyses()
    if df.empty:
        st.info("Belum ada riwayat tersimpan.")
        return
    st.dataframe(df, use_container_width=True)

    st.markdown("---")
    aid = st.selectbox("Pilih ID Analisis untuk membuka KKA lengkap", df["id"].tolist())
    if aid:
        an = fetch_analysis_by_id(aid)
        if an is not None:
            try:
                summary_data = json.loads(an["summary_json"]) if an["summary_json"] else {}
            except Exception:
                summary_data = {}
            entitas = summary_data.get("entitas", "-")
            periode = summary_data.get("periode", "-")
            score = summary_data.get("score", 0)
            raw_flagged = summary_data.get("flagged_data", {})
            flagged_data = {k: pd.DataFrame(v) if isinstance(v, list) else pd.DataFrame() for k, v in raw_flagged.items()}
            
            findings_df = fetch_findings(aid)
            findings_list = findings_df.to_dict("records")
            render_audit_results(findings_list, flagged_data, an["filename"], entitas, periode, score)

def page_landing():
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists(LOGO_FILENAME):
            st.image(LOGO_FILENAME, width=250)
        else:
            st.markdown("<h1 style='text-align:center; font-size: 5rem;'>F</h1>", unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #FFF0F4 0%, #FFF5E8 100%); border-radius: 20px; margin-bottom: 40px; border: 2px solid #FF5C8B;">
        <h1 style="color: #7A1F3D; font-size: 3rem; margin-bottom: 15px; font-weight: 800;">FINORA AI</h1>
        <p style="color: #2D1B24; font-size: 1.2rem; max-width: 800px; margin: 0 auto; line-height: 1.6;">
            Platform Kecerdasan Finansial untuk mendeteksi anomali, menyusun Kertas Kerja Audit (KKA), serta Analisis Laporan Keuangan & Cash Flow secara otomatis.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_cta1, col_cta2, col_cta3 = st.columns([1, 2, 1])
    with col_cta2:
        if st.button("🚀 MASUK KE APLIKASI UTAMA", use_container_width=True):
            st.session_state.app_started = True
            st.rerun()

# ============================================================
# MAIN ROUTER (PENGATUR NAVIGASI)
# ============================================================
def main():
    if "app_started" not in st.session_state:
        st.session_state.app_started = False

    if not st.session_state.app_started:
        page_landing()
    else:
        menu = st.sidebar.selectbox(
            "Pilih Menu Utama", 
            ["Dashboard", "Audit Profesional", "Analisis Laporan Keuangan", "Pusat Laporan & Arsip", "Kembali ke Beranda"]
        )
        
        if menu == "Dashboard":
            page_dashboard()
        elif menu == "Audit Profesional":
            page_audit()
        elif menu == "Analisis Laporan Keuangan":
            page_financial_analysis()
        elif menu == "Pusat Laporan & Arsip":
            page_simpan()
        elif menu == "Kembali ke Beranda":
            st.session_state.app_started = False
            st.rerun()

if __name__ == "__main__":
    main()
