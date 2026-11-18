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

        model_id = "gemini-3.8-flash"

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
# SMART FINANCIAL ENGINE
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
            return "<b>Kondisi:</b> Sedang. Masih cukup bergantung pada dana transfer pusat.<br><b>Penyebab:</b> Sektor unggulan daerah belum tergarap maksimal atau kepatuhan pajak masih rendah.<br><b>Dampak:</b> Rentan terhadap perubahan kebij
