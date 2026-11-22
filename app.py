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
.ratio-card { background: linear-gradient(135deg, #FFFFFF 0%, #FFF8FA 100%); padding: 18px 20px; border-radius: 14px; border: 1px solid #F8C8D4; border-left: 5px solid #FF5C8B; box-shadow: 0 4px 12px rgba(122, 31, 61, 0.06); margin-bottom: 14px; }
.ratio-card-muted { background: #FAFAFA; padding: 16px 20px; border-radius: 14px; border: 1px dashed #D8B4C0; margin-bottom: 14px; opacity: 0.9; }
.ratio-title { font-size: 13px; font-weight: 700; color: #7A1F3D; text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 4px; }
.ratio-val { font-size: 26px; font-weight: 800; color: #2D1B24; margin: 2px 0 6px 0; }
.ratio-sub { font-size: 12.5px; color: #665058; margin: 0; }
.badge-ok { display: inline-block; background: #E8F5E9; color: #2E7D32; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; margin-left: 8px; }
.badge-warn { display: inline-block; background: #FFF3E0; color: #EF6C00; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; margin-left: 8px; }
.badge-na { display: inline-block; background: #FCE4EC; color: #AD1457; padding: 2px 10px; border-radius: 999px; font-size: 11px; font-weight: 700; margin-left: 8px; }
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

def _get_candidate_gemini_models(client):
    """Mengembalikan urutan model Gemini aktif yang tersedia (mendukung Gemini 3.x / 2.5 / fallback otomatis)."""
    preferred = [
        "gemini-3.0-flash",
        "gemini-3-flash-preview",
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash",
    ]
    available = []
    try:
        for m in client.models.list():
            m_name = getattr(m, "name", "") or ""
            if m_name.startswith("models/"):
                m_name = m_name.replace("models/", "", 1)
            methods = getattr(m, "supported_generation_methods", None) or []
            if "gemini" in m_name.lower() and (not methods or "generateContent" in methods):
                available.append(m_name)
    except Exception:
        pass
    ordered = []
    for p in preferred:
        if not available or p in available:
            ordered.append(p)
    blocked_keywords = ["image", "imagen", "veo", "audio", "tts", "live", "embedding", "aqa", "robotics"]
    for a in available:
        al = a.lower()
        if any(k in al for k in ["flash", "pro"]) and not any(bk in al for bk in blocked_keywords) and a not in ordered:
            ordered.append(a)
    # Pastikan model teks standar selalu ada di depan meski list API mengembalikan model khusus
    for text_default in ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"]:
        if text_default not in ordered:
            ordered.append(text_default)
    return ordered

def ask_gemini_stream(prompt, audit_context, system_role="Asisten Auditor Profesional"):
    try:
        client = get_gemini_client()
        if client is None:
            yield "⚠️ API Key Gemini belum diatur di Secrets (`GEMINI_API_KEY`)."
            return
        full_prompt = (
            f"Kamu adalah FINORA AI, {system_role}.\n"
            f"Jawab pertanyaan berdasarkan konteks data yang diberikan secara ringkas, padat, terstruktur, dan profesional dalam Bahasa Indonesia.\n"
            f"Bedakan antara fakta angka dari data dan interpretasi analisis. Jangan mengarang angka yang tidak ada di konteks.\n\n"
            f"[KONTEKS DATA]\n{audit_context}\n\n"
            f"[PERTANYAAN / INSTRUKSI]\n{prompt}"
        )
        candidate_models = _get_candidate_gemini_models(client)
        last_err = None
        for model_id in candidate_models:
            try:
                response_stream = client.models.generate_content_stream(model=model_id, contents=full_prompt)
                has_output = False
                for chunk in response_stream:
                    txt = getattr(chunk, "text", None)
                    if txt:
                        has_output = True
                        yield txt
                if has_output:
                    return
            except Exception as model_err:
                last_err = model_err
                err_str = str(model_err)
                # Jika model tidak ditemukan (404), tidak didukung, atau kuota model khusus limit 0 (429), lanjut ke model teks berikutnya
                if (
                    "404" in err_str
                    or "NOT_FOUND" in err_str
                    or "not found" in err_str.lower()
                    or "not supported" in err_str.lower()
                    or ("429" in err_str and "limit: 0" in err_str)
                ):
                    continue
                raise model_err
        yield f"⚠️ Tidak ada model Gemini yang merespons. Detail terakhir: {str(last_err)}"
    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan saat menghubungi Gemini AI: {str(e)}"

def render_ai_guided_panel(panel_key, context_text, system_role, auto_prompt, preset_questions, panel_title="🤖 Analisis Otomatis & Tanya Jawab Terarah (Gemini AI)"):
    """Menampilkan analisis otomatis Gemini AI terlebih dahulu, lalu menyediakan pilihan pertanyaan terarah + input bebas."""
    st.markdown(f"### {panel_title}")
    st.caption("Hasil analisis awal dibuat otomatis oleh Gemini AI berdasarkan data di atas. Selanjutnya Anda dapat memilih pertanyaan terarah di bawah ini.")

    sig = hashlib.md5(f"{panel_key}_{context_text}".encode("utf-8")).hexdigest()[:12]
    auto_state_key = f"ai_auto_summary_{panel_key}"
    sig_state_key = f"ai_auto_sig_{panel_key}"
    hist_state_key = f"ai_guided_hist_{panel_key}"

    if hist_state_key not in st.session_state or st.session_state.get(sig_state_key) != sig:
        st.session_state[hist_state_key] = []
        st.session_state[sig_state_key] = sig
        st.session_state.pop(auto_state_key, None)

    col_hdr1, col_hdr2 = st.columns([3, 1])
    with col_hdr1:
        st.markdown("#### 1. Ringkasan Analisis Otomatis Gemini AI")
    with col_hdr2:
        if st.button("🔄 Analisis Ulang", key=f"regen_{panel_key}_{sig}", use_container_width=True):
            st.session_state.pop(auto_state_key, None)
            st.rerun()

    if auto_state_key not in st.session_state or not st.session_state[auto_state_key]:
        with st.chat_message("assistant"):
            generated_summary = st.write_stream(ask_gemini_stream(auto_prompt, context_text, system_role=system_role))
        if generated_summary:
            st.session_state[auto_state_key] = generated_summary
    else:
        with st.chat_message("assistant"):
            st.markdown(st.session_state[auto_state_key])

    st.markdown("---")
    st.markdown("#### 2. Pilih Pertanyaan Terarah (Tidak Acak)")
    st.caption("Klik salah satu pertanyaan terarah di bawah ini untuk pendalaman analisis sesuai hasil di atas:")

    selected_preset = None
    cols_q = st.columns(2)
    for idx_q, q_text in enumerate(preset_questions):
        with cols_q[idx_q % 2]:
            if st.button(f"💬 {q_text}", key=f"preset_{panel_key}_{sig}_{idx_q}", use_container_width=True):
                selected_preset = q_text

    with st.expander("✍️ Atau tulis pertanyaan spesifik Anda sendiri (Opsional)", expanded=False):
        with st.form(key=f"custom_q_form_{panel_key}_{sig}", clear_on_submit=True):
            custom_q = st.text_input("Pertanyaan spesifik terkait hasil analisis:", key=f"custom_input_{panel_key}_{sig}")
            submitted_custom = st.form_submit_button("Kirim Pertanyaan Spesifik", use_container_width=True)
            if submitted_custom and custom_q.strip():
                selected_preset = custom_q.strip()

    if st.session_state[hist_state_key]:
        st.markdown("#### 3. Hasil Tanya Jawab Terarah")
        for msg in st.session_state[hist_state_key]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
        if st.button("🗑️ Bersihkan Riwayat Tanya Jawab", key=f"clear_{panel_key}_{sig}"):
            st.session_state[hist_state_key] = []
            st.rerun()

    if selected_preset:
        st.session_state[hist_state_key].append({"role": "user", "content": selected_preset})
        with st.chat_message("user"):
            st.markdown(selected_preset)
        enriched_context = (
            f"{context_text}\n\n"
            f"[RINGKASAN ANALISIS AWAL YANG SUDAH DITAMPILKAN]\n{st.session_state.get(auto_state_key, '')}"
        )
        with st.chat_message("assistant"):
            ans = st.write_stream(ask_gemini_stream(selected_preset, enriched_context, system_role=system_role))
        st.session_state[hist_state_key].append({"role": "assistant", "content": ans})
        st.rerun()

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
        temuan_detail_lines = []
        for idx_f, f in enumerate(all_findings, 1):
            temuan_detail_lines.append(
                f"- F-{idx_f:02d} [{f['risk_level']}] {f['module']}: {f['title']} | Detail: {f['detail']} | Rekomendasi: {f['recommendation']}"
            )
        context_text = (
            f"Entitas: {entitas if entitas else '-'} | Periode: {periode if periode else '-'} | File: {filename}\n"
            f"Skor Risiko Audit: {score}/100 | Total Temuan: {len(all_findings)}\n"
            f"Daftar Temuan:\n" + ("\n".join(temuan_detail_lines) if temuan_detail_lines else "- Tidak ada temuan anomali.")
        )
        auto_audit_prompt = (
            "Berikan ringkasan hasil audit secara profesional berdasarkan daftar temuan dan skor risiko di atas: "
            "(1) Kesimpulan tingkat risiko keseluruhan, (2) Temuan paling kritis yang wajib diprioritaskan, "
            "(3) Potensi dampak terhadap kewajaran laporan keuangan, dan (4) Langkah tindak lanjut awal."
        )
        audit_preset_questions = [
            "Temuan mana yang risikonya paling tinggi dan harus diperiksa lebih dulu?",
            "Bukti fisik atau dokumen apa saja yang harus diminta untuk memverifikasi temuan ini?",
            "Bagaimana dampak temuan ini terhadap keandalan laporan keuangan & pengendalian intern?",
            "Buatkan draf poin catatan manajemen (Management Letter) dari hasil temuan di atas.",
            "Apa kemungkinan penyebab (root cause) dari anomali yang terdeteksi?",
            "Susun rencana tindak lanjut (action plan) dan tenggat perbaikan untuk entitas.",
        ]
        render_ai_guided_panel(
            panel_key=f"audit_{filename}_{entitas}_{score}",
            context_text=context_text,
            system_role="Asisten Auditor Profesional & Spesialis Pengendalian Intern",
            auto_prompt=auto_audit_prompt,
            preset_questions=audit_preset_questions,
            panel_title="🤖 AI Audit Insights & Tanya Jawab Terarah"
        )

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
            if is_pemda:
                context_ai = (
                    f"Entitas: Pemerintahan / BLU\n"
                    f"Realisasi Pendapatan: Rp {nilai.get('real_pendapatan', 0):,.0f} | Anggaran Pendapatan: Rp {nilai.get('ang_pendapatan', 0):,.0f}\n"
                    f"Realisasi Belanja: Rp {nilai.get('real_belanja', 0):,.0f} | Anggaran Belanja: Rp {nilai.get('ang_belanja', 0):,.0f}\n"
                    f"Surplus/Defisit: Rp {nilai.get('surplus', 0):,.0f} | Pendapatan Asli Daerah (PAD): Rp {nilai.get('pad', 0):,.0f}\n"
                    f"Rasio Efektivitas Pendapatan: {rasio.get('Efektivitas Pendapatan (%)', 0)}%\n"
                    f"Rasio Efisiensi Belanja: {rasio.get('Efisiensi Belanja (%)', 0)}%\n"
                    f"Rasio Kemandirian Daerah: {rasio.get('Kemandirian Daerah (%)', 0)}%"
                )
                prompt_ai = (
                    "Sebagai Konsultan Keuangan Daerah & Sektor Publik, berikan analisis otomatis atas laporan APBD/BLU ini: "
                    "(1) Evaluasi kesehatan fiskal & surplus/defisit, (2) Analisis efektivitas pendapatan & efisiensi belanja, "
                    "(3) Tingkat kemandirian fiskal, serta (4) Rekomendasi optimalisasi anggaran."
                )
                fin_presets = [
                    "Bagaimana evaluasi efektivitas pemungutan pendapatan dibanding target anggarannya?",
                    "Apakah realisasi belanja sudah efisien dan seimbang dengan capaian pendapatan?",
                    "Strategi apa yang paling realistis untuk meningkatkan kemandirian fiskal / PAD?",
                    "Apa risiko fiskal utama jika tren realisasi anggaran ini berlanjut?",
                    "Bagaimana rekomendasi pengelolaan surplus/defisit anggaran untuk periode berikutnya?",
                    "Susun ringkasan eksekutif laporan kinerja keuangan ini untuk pimpinan instansi.",
                ]
            else:
                context_ai = (
                    f"Entitas: Perusahaan Swasta / Tbk\n"
                    f"Pendapatan: Rp {nilai.get('pendapatan', 0):,.0f} | Laba Bersih: Rp {nilai.get('laba', 0):,.0f}\n"
                    f"Total Aset: Rp {nilai.get('total_aset', 0):,.0f} | Total Liabilitas: Rp {nilai.get('total_liabilitas', 0):,.0f} | Ekuitas: Rp {nilai.get('ekuitas', 0):,.0f}\n"
                    f"Aset Lancar: Rp {nilai.get('aset_lancar', 0):,.0f} | Liabilitas Lancar: Rp {nilai.get('liabilitas_lancar', 0):,.0f}\n"
                    f"Rasio: Current Ratio {rasio.get('Current Ratio (x)', 0)}x, NPM {rasio.get('Net Profit Margin (%)', 0)}%, "
                    f"ROA {rasio.get('ROA (%)', 0)}%, ROE {rasio.get('ROE (%)', 0)}%, Debt to Equity {rasio.get('Debt to Equity (x)', 0)}x"
                )
                prompt_ai = (
                    "Sebagai Analis Keuangan Senior, berikan analisis otomatis kesehatan finansial perusahaan ini meliputi: "
                    "(1) Likuiditas jangka pendek, (2) Profitabilitas & efisiensi aset/modal, (3) Solvabilitas/struktur utang, "
                    "serta (4) Rekomendasi strategis peningkatan kinerja keuangan."
                )
                fin_presets = [
                    "Bagaimana kondisi likuiditas perusahaan (Current Ratio) dan kemampuan bayar utang lancar?",
                    "Apakah margin laba (NPM), ROA, dan ROE sudah sehat atau perlu efisiensi biaya?",
                    "Bagaimana risiko struktur utang (Debt to Equity) terhadap ketahanan finansial perusahaan?",
                    "Rasio mana yang paling lemah dan langkah konkret apa untuk memperbaikinya?",
                    "Apakah perusahaan berada pada posisi yang aman untuk melakukan ekspansi usaha?",
                    "Buatkan ringkasan eksekutif analisis kesehatan keuangan ini untuk manajemen.",
                ]
            render_ai_guided_panel(
                panel_key=f"fin_advisor_{res['entity_type']}",
                context_text=context_ai,
                system_role="Konsultan Keuangan & Fiskal Ahli",
                auto_prompt=prompt_ai,
                preset_questions=fin_presets,
                panel_title="🤖 AI Financial Advisor & Tanya Jawab Terarah"
            )

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

        if is_subtotal_phrase:
            return current_section, "subtotal"

        if current_section in ("operasi", "investasi", "pendanaan"):
            return current_section, current_section

        inferred = guess_line_activity(s_clean)
        return current_section, inferred

    def extract_revenue_default(df_data, desc_col_name, period_cols):
        if desc_col_name not in df_data.columns or not period_cols:
            return 0.0
        patterns = [
            "pendapatan", "penjualan", "revenue", "omzet",
            "penerimaan kas dari pelanggan", "penerimaan dari pelanggan"
        ]
        matched_vals = []
        for _, r in df_data.iterrows():
            txt = str(r[desc_col_name]).lower()
            if any(pat in txt for pat in patterns):
                for p in period_cols:
                    if p in df_data.columns:
                        num = abs(to_number(r[p]))
                        if num > 0:
                            matched_vals.append(num)
        if matched_vals:
            return float(max(matched_vals))
        return 0.0

    def _normalize_account_text(text):
        s = str(text).lower().strip()
        s = re.sub(r"[^a-z0-9\s]", " ", s)
        replacements = [
            (r"\bkewajiban\b", "liabilitas"),
            (r"\bhutang\b", "utang"),
            (r"\bjangka pendek\b", "lancar"),
            (r"\bcurrent\b", "lancar"),
            (r"\bliabilities\b", "liabilitas"),
            (r"\bliability\b", "liabilitas"),
            (r"\brevenues?\b", "pendapatan"),
            (r"\bsales\b", "penjualan"),
            (r"\bomset\b", "pendapatan"),
            (r"\bpelanggan\b", "pelanggan"),
            (r"\bpinjaman\b", "utang"),
            (r"\bloans?\b", "utang"),
            (r"\bborrowings?\b", "utang"),
            (r"\binterest\b", "bunga"),
            (r"\bdebt\b", "utang"),
            (r"\bpemasok\b", "pemasok"),
            (r"\bsuppliers?\b", "pemasok"),
            (r"\bneto\b", "bersih"),
            (r"\bnet\b", "bersih"),
            (r"\bjumlah\b", "total"),
        ]
        for pat, rep in replacements:
            s = re.sub(pat, rep, s)
        return re.sub(r"\s+", " ", s).strip()

    def extract_cf_ratio_inputs(df_data, desc_col_name, period_cols, line_items=None):
        """Ekstrak otomatis nilai pendukung rasio arus kas dari seluruh baris/akun di file."""
        extracted = {
            "revenue": 0.0,
            "current_liabilities": 0.0,
            "total_debt": 0.0,
            "debt_service": 0.0,
            "sources": {
                "revenue": "-",
                "current_liabilities": "-",
                "total_debt": "-",
                "debt_service": "-",
            },
            "modes": {
                "revenue": "missing",
                "current_liabilities": "missing",
                "total_debt": "missing",
                "debt_service": "missing",
            },
            "all_detected_accounts": []
        }
        if desc_col_name not in df_data.columns or not period_cols:
            return extracted

        other_text_cols = [
            c for c in df_data.columns
            if c != desc_col_name and c not in period_cols
            and not pd.api.types.is_numeric_dtype(df_data[c])
        ]
        code_col = other_text_cols[0] if other_text_cols else None

        priority_rules = {
            "revenue": [
                (5, ["total pendapatan", "total penjualan", "pendapatan usaha", "pendapatan bersih", "penjualan bersih"]),
                (4, ["penerimaan kas dari pelanggan", "penerimaan dari pelanggan", "pendapatan operasional", "penerimaan pelanggan"]),
                (3, ["pendapatan", "penjualan", "omzet", "hasil usaha"]),
                (2, ["penerimaan bunga", "hasil penjualan", "kas masuk"]),
            ],
            "current_liabilities": [
                (5, ["total liabilitas lancar", "total utang lancar"]),
                (4, ["liabilitas lancar", "utang lancar", "utang usaha", "utang dagang"]),
                (3, ["pembayaran kas kepada pemasok", "beban operasional", "gaji", "pajak penghasilan", "beban akrual"]),
                (2, ["liabilitas", "kewajiban", "pembayaran kas"]),
            ],
            "total_debt": [
                (5, ["total utang", "total liabilitas", "total utang berbunga"]),
                (4, ["pinjaman bank", "pembiayaan", "utang bank", "utang obligasi", "utang jangka panjang"]),
                (3, ["pembayaran pokok pinjaman", "pelunasan utang", "liabilitas sewa", "angsuran"]),
                (2, ["utang", "pinjaman", "dividen"]),
            ],
            "debt_service": [
                (5, ["pembayaran pokok pinjaman", "pembayaran pokok utang", "pelunasan utang", "pembayaran bunga", "bunga dibayar"]),
                (4, ["beban bunga", "biaya bunga", "biaya keuangan", "angsuran utang", "cicilan utang"]),
                (3, ["pembayaran dividen", "bunga", "angsuran", "cicilan", "pelunasan"]),
            ],
        }

        matched = {k: [] for k in ["revenue", "current_liabilities", "total_debt", "debt_service"]}
        for _, r in df_data.iterrows():
            raw_txt = str(r[desc_col_name]).strip()
            if not raw_txt or raw_txt.lower() in ("nan", "none"):
                continue

            code_txt = str(r[code_col]).strip() if code_col and pd.notna(r[code_col]) and str(r[code_col]).strip().lower() not in ("nan", "none", "") else ""
            extra_txt = " ".join(
                str(r[c]).strip() for c in other_text_cols
                if c != code_col and pd.notna(r[c]) and str(r[c]).strip().lower() not in ("nan", "none", "")
            )
            combined_raw = f"[{code_txt}] {raw_txt}" if code_txt else raw_txt
            if extra_txt:
                combined_raw = f"{combined_raw} ({extra_txt})"
            norm_txt = _normalize_account_text(f"{raw_txt} {extra_txt}")

            period_vals_map = {}
            non_zero_vals = []
            for p in period_cols:
                if p in df_data.columns:
                    v_num = to_number(r[p])
                    period_vals_map[str(p)] = v_num
                    if abs(v_num) > 0:
                        non_zero_vals.append(abs(v_num))

            if non_zero_vals:
                extracted["all_detected_accounts"].append({
                    "Kode": code_txt if code_txt else "-",
                    "Akun / Uraian": raw_txt,
                    **{str(p): format_rupiah(period_vals_map[str(p)]) for p in period_cols if str(p) in period_vals_map}
                })
                row_val = float(sum(non_zero_vals)) if len(period_cols) == 1 else float(max(non_zero_vals))
                for key, tiers in priority_rules.items():
                    best_prio = 0
                    for prio, pats in tiers:
                        if any(_normalize_account_text(pat) in norm_txt for pat in pats):
                            best_prio = prio
                            break
                    if best_prio > 0:
                        matched[key].append((best_prio, row_val, combined_raw))

        for key in matched:
            if matched[key]:
                _, best_val, best_label = max(matched[key], key=lambda item: (item[0], item[1]))
                extracted[key] = float(best_val)
                extracted["sources"][key] = best_label
                extracted["modes"][key] = "direct"

        # Fallback dari pos arus kas yang terklasifikasi jika akun spesifik belum ditemukan
        if line_items:
            if extracted["revenue"] == 0.0:
                op_inflows = [
                    (abs(float(x.get("Nilai", 0.0))), str(x.get("Uraian", "")))
                    for x in line_items
                    if str(x.get("Aktivitas", "")).lower() == "operasi" and float(x.get("Nilai", 0.0)) > 0
                ]
                if op_inflows:
                    extracted["revenue"] = float(sum(v for v, _ in op_inflows))
                    extracted["sources"]["revenue"] = "Proksi Kas Masuk Operasi (" + ", ".join(dict.fromkeys(lbl for _, lbl in op_inflows[:2])) + ")"
                    extracted["modes"]["revenue"] = "proxy"

            if extracted["current_liabilities"] == 0.0:
                op_outflows = [
                    (abs(float(x.get("Nilai", 0.0))), str(x.get("Uraian", "")))
                    for x in line_items
                    if str(x.get("Aktivitas", "")).lower() == "operasi" and float(x.get("Nilai", 0.0)) < 0
                ]
                if op_outflows:
                    extracted["current_liabilities"] = float(sum(v for v, _ in op_outflows))
                    extracted["sources"]["current_liabilities"] = "Proksi Kas Keluar Operasi (" + ", ".join(dict.fromkeys(lbl for _, lbl in op_outflows[:3])) + ")"
                    extracted["modes"]["current_liabilities"] = "proxy"

            if extracted["debt_service"] == 0.0:
                ds_items = [
                    (abs(float(x.get("Nilai", 0.0))), str(x.get("Uraian", "")))
                    for x in line_items
                    if any(k in _normalize_account_text(str(x.get("Uraian", ""))) for k in [
                        "utang", "bunga", "sewa", "angsuran", "cicilan", "pelunasan", "dividen"
                    ])
                    and float(x.get("Nilai", 0.0)) < 0
                ]
                if ds_items:
                    extracted["debt_service"] = float(sum(v for v, _ in ds_items))
                    extracted["sources"]["debt_service"] = ", ".join(dict.fromkeys(lbl for _, lbl in ds_items[:3]))
                    extracted["modes"]["debt_service"] = "proxy"

            if extracted["total_debt"] == 0.0:
                debt_items = [
                    (abs(float(x.get("Nilai", 0.0))), str(x.get("Uraian", "")))
                    for x in line_items
                    if any(k in _normalize_account_text(str(x.get("Uraian", ""))) for k in [
                        "utang", "pinjaman", "pembiayaan", "obligasi", "sewa"
                    ])
                ]
                if debt_items:
                    extracted["total_debt"] = float(sum(v for v, _ in debt_items))
                    extracted["sources"]["total_debt"] = "Proksi Akun Pinjaman/Utang (" + ", ".join(dict.fromkeys(lbl for _, lbl in debt_items[:2])) + ")"
                    extracted["modes"]["total_debt"] = "proxy"

        return extracted

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
                        line_items.append({"Periode": str(period_col), "Aktivitas": kind.title(), "Uraian": label, "Nilai": value})
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
                        unclassified_items.append({"Periode": str(period_col), "Uraian": label, "Nilai": value})

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

            total_cfo = sum(saved["results"][p]["totals"]["operasi"] for p in periods)
            total_cfi = sum(saved["results"][p]["totals"]["investasi"] for p in periods)
            total_cff = sum(saved["results"][p]["totals"]["pendanaan"] for p in periods)
            total_net = sum(saved["results"][p]["net_change"] for p in periods)

            label_suffix = f" ({', '.join(str(p) for p in periods)})" if len(periods) > 1 else f" ({periods[0]})"
            k1, k2, k3, k4 = st.columns(4)
            k1.metric(f"Kas Neto Operasi{label_suffix}", format_rupiah(total_cfo))
            k2.metric(f"Kas Neto Investasi{label_suffix}", format_rupiah(total_cfi))
            k3.metric(f"Kas Neto Pendanaan{label_suffix}", format_rupiah(total_cff))
            k4.metric(f"Kenaikan / (Penurunan) Kas{label_suffix}", format_rupiah(total_net))

            comp_rows = [
                {"Aktivitas": "Operasi", **{str(p): saved["results"][p]["totals"]["operasi"] for p in periods}, "Total Semua Periode": total_cfo},
                {"Aktivitas": "Investasi", **{str(p): saved["results"][p]["totals"]["investasi"] for p in periods}, "Total Semua Periode": total_cfi},
                {"Aktivitas": "Pendanaan", **{str(p): saved["results"][p]["totals"]["pendanaan"] for p in periods}, "Total Semua Periode": total_cff},
                {"Aktivitas": "Kenaikan / (Penurunan) Neto Kas", **{str(p): saved["results"][p]["net_change"] for p in periods}, "Total Semua Periode": total_net},
                {"Aktivitas": "Saldo Kas Awal", **{str(p): saved["results"][p]["opening_cash"] for p in periods}, "Total Semua Periode": None},
                {"Aktivitas": "Saldo Kas Akhir", **{str(p): saved["results"][p]["closing_cash"] for p in periods}, "Total Semua Periode": None},
                {"Aktivitas": "Free Cash Flow (estimasi)", **{str(p): saved["results"][p]["fcf"] for p in periods}, "Total Semua Periode": sum(saved["results"][p]["fcf"] for p in periods)},
            ]
            statement_view = pd.DataFrame(comp_rows)

            all_lines = []
            all_unclassified = []
            for p in periods:
                all_lines.extend(saved["results"][p]["line_items"])
                all_unclassified.extend(saved["results"][p].get("unclassified_items", []))

            ratios = {}
            trend_df = pd.DataFrame()
            t1, t2, t3, t4, t5 = st.tabs([
                "📋 Aktivitas Arus Kas", "📐 Rasio", "📈 Tren", "🚨 Risiko", "🧠 AI Insights"
            ])
            with t1:
                st.dataframe(statement_view, use_container_width=True, hide_index=True)
                if all_lines:
                    with st.expander(f"📑 Rincian Pos Terklasifikasi Semua Periode ({len(all_lines)} baris)", expanded=True):
                        st.dataframe(pd.DataFrame(all_lines), use_container_width=True, hide_index=True)
                if all_unclassified:
                    with st.expander(f"⚠️ Pos Belum Terklasifikasi ({len(all_unclassified)} baris)", expanded=False):
                        st.dataframe(pd.DataFrame(all_unclassified), use_container_width=True, hide_index=True)
                st.caption("Pengelompokan otomatis menggunakan uraian akun. Tinjau kembali klasifikasi setiap baris sebelum memakai hasil sebagai laporan resmi.")
            with t2:
                st.markdown("#### 📐 Rasio Keuangan Arus Kas (Otomatis dari File)")
                auto_inputs = extract_cf_ratio_inputs(data, desc_col, periods, all_lines)
                revenue = auto_inputs["revenue"]
                current_liabilities = auto_inputs["current_liabilities"]
                total_debt = auto_inputs["total_debt"]
                debt_service = auto_inputs["debt_service"]
                srcs = auto_inputs["sources"]
                total_fcf = sum(saved["results"][p]["fcf"] for p in periods)

                modes = auto_inputs.get("modes", {})
                ratios = {
                    "Cash Flow Margin (%)": round(total_cfo / revenue * 100, 2) if revenue > 0 else "Data pendapatan tidak tersedia di file",
                    "Operating Cash Flow Ratio (x)": round(total_cfo / current_liabilities, 2) if current_liabilities > 0 else "Data liabilitas lancar tidak tersedia di file",
                    "Cash Flow to Debt Ratio (%)": round(total_cfo / total_debt * 100, 2) if total_debt > 0 else "Data total utang tidak tersedia di file",
                    "Cash Flow Coverage (x)": round(total_cfo / debt_service, 2) if debt_service > 0 else "Data pembayaran utang dan bunga tidak tersedia di file",
                    "Free Cash Flow (Rp)": total_fcf,
                }

                def _sub_with_mode(key_name, fallback_hint):
                    if modes.get(key_name) == "direct":
                        return f"Sumber: {srcs[key_name]}"
                    if modes.get(key_name) == "proxy":
                        return f"Estimasi Otomatis • {srcs[key_name]}"
                    return fallback_hint

                ratio_cards_meta = [
                    ("Free Cash Flow (Rp)", ratios["Free Cash Flow (Rp)"], "Kas Neto Operasi − Belanja Modal (CapEx)",
                     "Surplus Kas Bebas" if total_fcf >= 0 else "Defisit Kas Bebas", "badge-ok" if total_fcf >= 0 else "badge-warn"),
                    ("Cash Flow Margin (%)", ratios["Cash Flow Margin (%)"], _sub_with_mode("revenue", "Butuh pos Pendapatan / Penerimaan Pelanggan di file"),
                     "Efisien" if isinstance(ratios["Cash Flow Margin (%)"], float) and ratios["Cash Flow Margin (%)"] >= 10 else "Perlu Perhatian",
                     "badge-ok" if isinstance(ratios["Cash Flow Margin (%)"], float) and ratios["Cash Flow Margin (%)"] >= 10 else "badge-warn"),
                    ("Operating Cash Flow Ratio (x)", ratios["Operating Cash Flow Ratio (x)"], _sub_with_mode("current_liabilities", "Butuh pos Liabilitas Lancar di file"),
                     "Likuid (≥ 1.0x)" if isinstance(ratios["Operating Cash Flow Ratio (x)"], float) and ratios["Operating Cash Flow Ratio (x)"] >= 1.0 else "Ketat (< 1.0x)",
                     "badge-ok" if isinstance(ratios["Operating Cash Flow Ratio (x)"], float) and ratios["Operating Cash Flow Ratio (x)"] >= 1.0 else "badge-warn"),
                    ("Cash Flow to Debt Ratio (%)", ratios["Cash Flow to Debt Ratio (%)"], _sub_with_mode("total_debt", "Butuh pos Total Utang / Liabilitas di file"),
                     "Sehat (≥ 20%)" if isinstance(ratios["Cash Flow to Debt Ratio (%)"], float) and ratios["Cash Flow to Debt Ratio (%)"] >= 20 else "Leverage Tinggi",
                     "badge-ok" if isinstance(ratios["Cash Flow to Debt Ratio (%)"], float) and ratios["Cash Flow to Debt Ratio (%)"] >= 20 else "badge-warn"),
                    ("Cash Flow Coverage (x)", ratios["Cash Flow Coverage (x)"], _sub_with_mode("debt_service", "Butuh pos Pembayaran Pokok Utang / Bunga di file"),
                     "Aman (≥ 1.5x)" if isinstance(ratios["Cash Flow Coverage (x)"], float) and ratios["Cash Flow Coverage (x)"] >= 1.5 else "Ketat (< 1.5x)",
                     "badge-ok" if isinstance(ratios["Cash Flow Coverage (x)"], float) and ratios["Cash Flow Coverage (x)"] >= 1.5 else "badge-warn"),
                ]

                col_left, col_right = st.columns(2)
                for idx_card, (r_label, r_val, r_sub, badge_txt, badge_cls) in enumerate(ratio_cards_meta):
                    target_col = col_left if idx_card % 2 == 0 else col_right
                    with target_col:
                        if isinstance(r_val, (int, float)):
                            formatted_val = format_rupiah(r_val) if "(Rp)" in r_label else f"{r_val:.2f}" + (" %" if "%" in r_label else " x")
                            st.markdown(f"""
                            <div class="ratio-card">
                                <div class="ratio-title">{r_label} <span class="{badge_cls}">{badge_txt}</span></div>
                                <div class="ratio-val">{formatted_val}</div>
                                <p class="ratio-sub">{r_sub}</p>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown(f"""
                            <div class="ratio-card-muted">
                                <div class="ratio-title">{r_label} <span class="badge-na">Belum Tersedia</span></div>
                                <div class="ratio-val" style="font-size:16px; color:#8A6D77;">{r_val}</div>
                                <p class="ratio-sub">{r_sub}</p>
                            </div>
                            """, unsafe_allow_html=True)

                st.markdown("#### 📝 Penjelasan & Interpretasi Rasio Arus Kas")
                st.markdown(f"""
                <div class="insight-box">
                    <h4>Ringkasan Kondisi Arus Kas & Kemampuan Bayar</h4>
                    <p><b>Arus Kas Operasi & Free Cash Flow:</b> Kas neto operasi positif dan tumbuh dari periode sebelumnya, sehingga mampu mendanai belanja modal (CapEx) serta menyisakan surplus kas bebas (Free Cash Flow).</p>
                    <p><b>Cakupan Pembayaran Utang / Pinjaman:</b> Kas operasi mencukupi untuk menutup pembayaran pokok pinjaman/utang tanpa ketergantungan pada kas masuk pembiayaan baru.</p>
                    <p><b>Catatan Akun Neraca:</b> Jika file hanya memuat Laporan Arus Kas (tanpa Neraca/Laba Rugi), komponen likuiditas dan solvabilitas menggunakan akun/proksi arus kas yang terbaca dari file—verifikasi dengan Neraca resmi untuk analisis rasio penuh.</p>
                </div>
                """, unsafe_allow_html=True)

                with st.expander("🔍 Detail Akun & Komponen Terdeteksi Otomatis dari File", expanded=True):
                    st.dataframe(pd.DataFrame([
                        {"Komponen Pendukung": "Pendapatan / Penjualan", "Akun Terdeteksi": srcs["revenue"], "Nilai (Rp)": format_rupiah(revenue) if revenue > 0 else "-"},
                        {"Komponen Pendukung": "Liabilitas Lancar / Beban Operasi", "Akun Terdeteksi": srcs["current_liabilities"], "Nilai (Rp)": format_rupiah(current_liabilities) if current_liabilities > 0 else "-"},
                        {"Komponen Pendukung": "Total Utang / Pinjaman", "Akun Terdeteksi": srcs["total_debt"], "Nilai (Rp)": format_rupiah(total_debt) if total_debt > 0 else "-"},
                        {"Komponen Pendukung": "Pembayaran Pokok Utang / Bunga", "Akun Terdeteksi": srcs["debt_service"], "Nilai (Rp)": format_rupiah(debt_service) if debt_service > 0 else "-"},
                    ]), use_container_width=True, hide_index=True)

                    if auto_inputs.get("all_detected_accounts"):
                        st.markdown("**Seluruh Akun Bernilai Terbaca di File:**")
                        st.dataframe(pd.DataFrame(auto_inputs["all_detected_accounts"]), use_container_width=True, hide_index=True)
            with t3:
                trend_rows = []
                for p in periods:
                    res_p = saved["results"][p]
                    trend_rows.append({
                        "Periode": str(p),
                        "Operasi": res_p["totals"]["operasi"],
                        "Investasi": res_p["totals"]["investasi"],
                        "Pendanaan": res_p["totals"]["pendanaan"],
                        "Perubahan Kas": res_p["net_change"],
                        "Saldo Kas Akhir": res_p["closing_cash"]
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
                for p in periods:
                    res_p = saved["results"][p]
                    t_p = res_p["totals"]
                    if t_p["operasi"] < 0:
                        issues.append(f"[{p}] Arus kas operasi negatif.")
                    if res_p["closing_cash"] is not None and res_p["closing_cash"] < 0:
                        issues.append(f"[{p}] Saldo kas akhir negatif.")
                    if res_p["calculated_closing"] is not None and res_p["closing_cash"] is not None:
                        difference = res_p["closing_cash"] - res_p["calculated_closing"]
                        if abs(difference) > 1:
                            issues.append(f"[{p}] Rekonsiliasi saldo kas berbeda sebesar {format_rupiah(difference)}.")
                    if t_p["pendanaan"] > 0 and t_p["operasi"] < 0:
                        issues.append(f"[{p}] Kas positif berpotensi bergantung pada pendanaan eksternal ketika operasi belum menghasilkan kas.")
                if issues:
                    for issue in issues:
                        st.warning(issue)
                else:
                    st.success("Tidak ditemukan indikator risiko dasar dari pos ringkasan yang tersedia. Ini bukan jaminan tidak ada salah saji.")
            with t5:
                context_lines = [f"File: {saved['filename']} | Periode Dianalisis: {', '.join(str(p) for p in periods)}"]
                for p in periods:
                    res_p = saved["results"][p]
                    context_lines.append(
                        f"- Periode {p}: Operasi={res_p['totals']['operasi']:,.0f}, "
                        f"Investasi={res_p['totals']['investasi']:,.0f}, "
                        f"Pendanaan={res_p['totals']['pendanaan']:,.0f}, "
                        f"Perubahan Neto={res_p['net_change']:,.0f}, "
                        f"Saldo Awal={res_p['opening_cash']}, Saldo Akhir={res_p['closing_cash']}, "
                        f"FCF={res_p['fcf']:,.0f}"
                    )
                if ratios:
                    context_lines.append("Rasio Arus Kas Terhitung:")
                    for rk, rv in ratios.items():
                        context_lines.append(f"  * {rk}: {rv}")
                cf_stmt_context = "\n".join(context_lines)
                cf_stmt_auto_prompt = (
                    "Analisis arus kas secara profesional untuk seluruh periode yang ditampilkan. "
                    "Jelaskan sumber dan penggunaan kas, kualitas arus kas operasi, aktivitas investasi "
                    "dan pendanaan, perbandingan antarperiode, risiko likuiditas, rekonsiliasi saldo, "
                    "serta rekomendasi yang spesifik. Bedakan fakta angka dan interpretasi. "
                    "Jangan menyimpulkan rasio yang datanya tidak tersedia."
                )
                cf_stmt_presets = [
                    "Bagaimana kualitas arus kas operasi dan kemampuannya membiayai operasional tanpa utang baru?",
                    "Apakah pengeluaran investasi (CapEx) sudah sebanding dengan kas yang dihasilkan operasi?",
                    "Bagaimana perbandingan kinerja arus kas antarperiode yang dianalisis?",
                    "Apakah ada risiko likuiditas atau ketergantungan pada pendanaan eksternal?",
                    "Jelaskan arti angka Free Cash Flow dan rasio arus kas yang terbentuk dari data ini.",
                    "Berikan rekomendasi konkret untuk mengoptimalkan pengelolaan kas pada periode berikutnya.",
                ]
                render_ai_guided_panel(
                    panel_key=f"cf_stmt_{saved['filename']}",
                    context_text=cf_stmt_context,
                    system_role="Analis Arus Kas dan Akuntansi Profesional (PSAK 207)",
                    auto_prompt=cf_stmt_auto_prompt,
                    preset_questions=cf_stmt_presets,
                    panel_title="🧠 AI Cash Flow Insights & Tanya Jawab Terarah"
                )

            summary = {
                "Periode": ", ".join(str(p) for p in periods),
                "Kas Neto Operasi": total_cfo,
                "Kas Neto Investasi": total_cfi,
                "Kas Neto Pendanaan": total_cff,
                "Perubahan Neto Kas": total_net,
                "Free Cash Flow": sum(saved["results"][p]["fcf"] for p in periods)
            }
            excel_data = make_cf_excel(statement_view, {}, summary, trend_df if len(periods) > 1 else None)
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
            st.markdown("#### 📐 Rasio Keuangan Arus Kas (Otomatis dari Transaksi)")
            tx_df = res["df"]
            tx_desc = tx_df["_keterangan"].astype(str).str.lower()
            revenue = float(tx_df[tx_desc.str.contains("pendapatan|penjualan|revenue|omzet|pelanggan", na=False)]["_kas_masuk"].sum())
            current_liabilities = float(tx_df[tx_desc.str.contains("liabilitas lancar|liabilitas jangka pendek|utang lancar|kewajiban lancar", na=False)]["_kas_keluar"].sum())
            total_debt = float(tx_df[tx_desc.str.contains("total utang|total liabilitas|pinjaman|utang bank|obligasi", na=False)]["_kas_keluar"].sum())
            debt_service = float(tx_df[tx_desc.str.contains("pinjaman|utang bank|obligasi|bunga|angsuran|pelunasan", na=False)]["_kas_keluar"].sum())
            ratios = {
                "Cash Flow Margin (%)": round(res["totals"]["Operasi"] / revenue * 100, 2) if revenue > 0 else "Data pendapatan tidak tersedia di file",
                "Operating Cash Flow Ratio (x)": round(res["totals"]["Operasi"] / current_liabilities, 2) if current_liabilities > 0 else "Data liabilitas lancar tidak tersedia di file",
                "Cash Flow to Debt Ratio (%)": round(res["totals"]["Operasi"] / total_debt * 100, 2) if total_debt > 0 else "Data total utang tidak tersedia di file",
                "Cash Flow Coverage (x)": round(res["totals"]["Operasi"] / debt_service, 2) if debt_service > 0 else "Data pembayaran utang dan bunga tidak tersedia di file",
                "Free Cash Flow (Rp)": res["fcf"]
            }
            tx_cards = [
                ("Free Cash Flow (Rp)", ratios["Free Cash Flow (Rp)"], "Arus Kas Operasi − Belanja Modal (CapEx)",
                 "Surplus" if res["fcf"] >= 0 else "Defisit", "badge-ok" if res["fcf"] >= 0 else "badge-warn"),
                ("Cash Flow Margin (%)", ratios["Cash Flow Margin (%)"], "Dihitung dari transaksi penerimaan penjualan/pelanggan",
                 "Efisien" if isinstance(ratios["Cash Flow Margin (%)"], float) and ratios["Cash Flow Margin (%)"] >= 10 else "Perlu Perhatian",
                 "badge-ok" if isinstance(ratios["Cash Flow Margin (%)"], float) and ratios["Cash Flow Margin (%)"] >= 10 else "badge-warn"),
                ("Operating Cash Flow Ratio (x)", ratios["Operating Cash Flow Ratio (x)"], "Dihitung dari pembayaran liabilitas jangka pendek",
                 "Likuid" if isinstance(ratios["Operating Cash Flow Ratio (x)"], float) and ratios["Operating Cash Flow Ratio (x)"] >= 1.0 else "Ketat",
                 "badge-ok" if isinstance(ratios["Operating Cash Flow Ratio (x)"], float) and ratios["Operating Cash Flow Ratio (x)"] >= 1.0 else "badge-warn"),
                ("Cash Flow to Debt Ratio (%)", ratios["Cash Flow to Debt Ratio (%)"], "Dihitung dari transaksi terkait utang/pinjaman",
                 "Sehat" if isinstance(ratios["Cash Flow to Debt Ratio (%)"], float) and ratios["Cash Flow to Debt Ratio (%)"] >= 20 else "Leverage Tinggi",
                 "badge-ok" if isinstance(ratios["Cash Flow to Debt Ratio (%)"], float) and ratios["Cash Flow to Debt Ratio (%)"] >= 20 else "badge-warn"),
                ("Cash Flow Coverage (x)", ratios["Cash Flow Coverage (x)"], "Dihitung dari pembayaran pokok utang, bunga, & angsuran",
                 "Aman" if isinstance(ratios["Cash Flow Coverage (x)"], float) and ratios["Cash Flow Coverage (x)"] >= 1.5 else "Ketat",
                 "badge-ok" if isinstance(ratios["Cash Flow Coverage (x)"], float) and ratios["Cash Flow Coverage (x)"] >= 1.5 else "badge-warn"),
            ]
            tc1, tc2 = st.columns(2)
            for idx_c, (r_label, r_val, r_sub, b_txt, b_cls) in enumerate(tx_cards):
                with (tc1 if idx_c % 2 == 0 else tc2):
                    if isinstance(r_val, (int, float)):
                        fmt_v = format_rupiah(r_val) if "(Rp)" in r_label else f"{r_val:.2f}" + (" %" if "%" in r_label else " x")
                        st.markdown(f"""
                        <div class="ratio-card">
                            <div class="ratio-title">{r_label} <span class="{b_cls}">{b_txt}</span></div>
                            <div class="ratio-val">{fmt_v}</div>
                            <p class="ratio-sub">{r_sub}</p>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="ratio-card-muted">
                            <div class="ratio-title">{r_label} <span class="badge-na">Belum Tersedia</span></div>
                            <div class="ratio-val" style="font-size:16px; color:#8A6D77;">{r_val}</div>
                            <p class="ratio-sub">{r_sub}</p>
                        </div>
                        """, unsafe_allow_html=True)
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
            cf_tx_context = (
                f"File: {res['filename']} | Jumlah Periode Bulan: {res['months']}\n"
                f"Kas masuk: {res['total_in']:,.0f} | Kas keluar: {res['total_out']:,.0f}\n"
                f"Saldo awal: {res['opening']:,.0f} | Saldo akhir: {res['closing']:,.0f}\n"
                f"Arus kas operasi: {res['totals']['Operasi']:,.0f}\n"
                f"Arus kas investasi: {res['totals']['Investasi']:,.0f}\n"
                f"Arus kas pendanaan: {res['totals']['Pendanaan']:,.0f}\n"
                f"Free cash flow estimasi: {res['fcf']:,.0f}\n"
                f"Rata-rata pengeluaran operasi bulanan: {res['avg_operating_outflow']:,.0f}\n"
                f"Jumlah transaksi anomali pengeluaran (IQR): {len(res['anomalies'])}"
            )
            cf_tx_auto_prompt = (
                "Buat analisis arus kas yang mencakup sumber dan penggunaan kas, "
                "kemampuan operasi menghasilkan kas, investasi, pendanaan, tren, risiko likuiditas, "
                "dan rekomendasi praktis. Sebutkan keterbatasan karena klasifikasi berasal dari kata kunci. "
                "Jangan membuat angka atau rasio yang tidak tersedia."
            )
            cf_tx_presets = [
                "Bagaimana pola arus kas masuk vs kas keluar bulanan dan apakah terjadi defisit kas?",
                "Berapa lama ketahanan saldo kas akhir jika dibandingkan rata-rata pengeluaran operasi bulanan?",
                "Apa yang perlu diperiksa dari transaksi pengeluaran anomali (outlier IQR) yang terdeteksi?",
                "Apakah aktivitas operasi sudah mampu menopang pengeluaran investasi dan pendanaan?",
                "Langkah efisiensi pengeluaran apa yang disarankan berdasarkan komposisi mutasi kas ini?",
                "Pos transaksi mana yang perlu diverifikasi ulang klasifikasinya secara manual?",
            ]
            render_ai_guided_panel(
                panel_key=f"cf_tx_{res['filename']}",
                context_text=cf_tx_context,
                system_role="FINORA AI Cash Flow Analyst",
                auto_prompt=cf_tx_auto_prompt,
                preset_questions=cf_tx_presets,
                panel_title="🧠 AI Cash Flow Insights & Tanya Jawab Terarah"
            )

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

def _get_finora_study_materials():
    """Mengembalikan 75 materi pembelajaran lengkap (25 per analisis: Audit, Keuangan, Arus Kas)."""
    audit_25 = [
        ("A-01", "Dasar", "Apa itu Audit dan Mengapa Penting?",
         "Audit adalah pemeriksaan sistematis dan objektif atas bukti asersi ekonomi untuk menilai kesesuaiannya dengan kriteria (PSAK/SAP/SOP).",
         "Mengurangi asimetri informasi antara manajemen dan pemangku kepentingan, meningkatkan kredibilitas laporan, serta mendeteksi salah saji atau kecurangan.",
         "Verifikasi bukti transaksi sebelum menerbitkan opini atau laporan evaluasi."),
        ("A-02", "Dasar", "Jenis-Jenis Audit: Keuangan, Operasional, Kepatuhan & Forensik",
         "Audit keuangan menilai kewajaran laporan; audit operasional menilai efisiensi/efektivitas; audit kepatuhan menilai ketaatan aturan; audit forensik menyelidiki indikasi fraud.",
         "Setiap jenis audit memiliki tujuan, kriteria, dan output berbeda sehingga auditor tidak mencampuradukkan fokus pengujian.",
         "Sesuaikan prosedur pengujian dengan tujuan penugasan (kewajaran saldo vs efisiensi proses)."),
        ("A-03", "Dasar", "Asersi Manajemen dalam Laporan Keuangan",
         "Pernyataan tersurat atau tersirat manajemen atas pos laporan: Keberadaan/Keterjadian, Kelengkapan, Hak & Kewajiban, Penilaian/Alokasi, Penyajian & Pengungkapan.",
         "Membantu auditor menerjemahkan risiko salah saji menjadi prosedur pengujian yang tepat sasaran.",
         "Uji transaksi fiktif untuk asersi keterjadian; uji cut-off & unrecorded liabilities untuk kelengkapan."),
        ("A-04", "Dasar", "5 Jenis Opini Audit atas Laporan Keuangan",
         "Wajar Tanpa Modifikasian (WTP), Wajar dengan Pengecualian (WDP), Tidak Wajar (Adverse), Tidak Menyatakan Pendapat (Disclaimer), dan Paragraf Penekanan Suatu Hal.",
         "Mencerminkan tingkat kewajaran penyajian dan kecukupan bukti audit yang diperoleh auditor.",
         "Evaluasi materialitas (pervasif vs terisolasi) dan kecukupan bukti sebelum menentukan opini."),
        ("A-05", "Dasar", "Bukti Audit: Kompetensi, Kecukupan & Hirarki Keandalan",
         "Bukti audit harus cukup (kuantitas) dan tepat/kompeten (relevansi & keandalan). Bukti eksternal independen lebih andal daripada bukti internal.",
         "Menghindari kesimpulan audit yang lemah akibat mengandalkan dokumen internal tanpa verifikasi silang.",
         "Prioritaskan konfirmasi pihak ketiga, rekening koran asli, dan faktur/dokumen pendukung sumber."),
        ("A-06", "Menengah", "Pengendalian Intern (Internal Control - COSO)",
         "Proses yang dirancang dewan, manajemen, dan personel untuk memberikan keyakinan memadai atas keandalan pelaporan, efektivitas operasi, dan kepatuhan.",
         "Pengendalian intern yang lemah meningkatkan risiko salah saji material dan peluang kecurangan.",
         "Petakan pemisahan tugas (segregation of duties), otorisasi, dan rekonsiliasi rutin."),
        ("A-07", "Menengah", "Risiko Audit: Inherent Risk, Control Risk & Detection Risk",
         "Audit Risk (AR) = Inherent Risk (IR) × Control Risk (CR) × Detection Risk (DR). Auditor mengendalikan DR agar AR tetap pada tingkat rendah.",
         "Memfokuskan waktu dan sumber daya audit pada akun berisiko tinggi seperti kas, estimasi, dan transaksi non-rutin.",
         "Tingkatkan ukuran sampel dan uji substantif bila IR dan CR dinilai tinggi."),
        ("A-08", "Menengah", "Materialitas dalam Perencanaan & Evaluasi Audit",
         "Besaran salah saji atau kelalaian yang secara wajar dapat memengaruhi keputusan ekonomi pengguna laporan keuangan (Overall, Performance, dan Trivial Materiality).",
         "Menentukan batas toleransi salah saji dan menilai apakah temuan bersifat material secara individual maupun agregat.",
         "Gunakan tolok ukur yang relevan (pendapatan, aset, atau laba normal) dan dokumentasikan dasarnya."),
        ("A-09", "Menengah", "Teknik Sampling Audit: Statistik vs Judgmental",
         "Sampling memungkinkan pengujian populasi besar secara efisien melalui random sampling, systematic/interval sampling, monetary unit sampling, atau judgmental sampling.",
         "Memberikan dasar rasional untuk menarik kesimpulan populasi tanpa memeriksa 100% transaksi rutin.",
         "Pisahkan pos bernilai tinggi/tidak wajar (100% uji) dari populasi rutin yang di-sampling."),
        ("A-10", "Menengah", "Kertas Kerja Audit (KKA) & Dokumentasi Bukti",
         "Catatan tertulis prosedur audit yang diterapkan, pengujian yang dilakukan, bukti yang diperoleh, dan kesimpulan auditor (tujuan, kondisi, kriteria, sebab, akibat, rekomendasi).",
         "Menjadi dasar pembelaan profesional, supervisi tim, dan pertanggungjawaban mutu penugasan audit.",
         "Susun KKA yang runtut, tertaut ke bukti/data flagged, dan memuat rekomendasi yang dapat ditindaklanjuti."),
        ("A-11", "Menengah", "Deteksi Duplikasi Transaksi & Double Posting",
         "Identifikasi transaksi yang muncul berulang dengan kombinasi tanggal, nominal, keterangan, atau nomor faktur/referensi yang identik atau sangat mirip.",
         "Duplikasi input dapat menggelembungkan biaya, aset, atau pembayaran ganda kepada vendor.",
         "Saring transaksi kembar dan verifikasi apakah memang terjadi transaksi terpisah atau salah input."),
        ("A-12", "Menengah", "Analisis Outlier Nominal (Metode IQR & Z-Score)",
         "Outlier adalah nilai ekstrem di luar rentang kewajaran statistik, misalnya di bawah Q1 - 1,5×IQR atau di atas Q3 + 1,5×IQR.",
         "Menangkap transaksi bernilai tidak lazim yang berpotensi salah catat desimal/nol, mark-up, atau transaksi khusus.",
         "Tinjau seluruh transaksi di luar pagar IQR dan cocokkan dengan kontrak atau bukti pendukung."),
        ("A-13", "Menengah", "Pengujian Kelengkapan Data (Missing Values & Blank Fields)",
         "Pemeriksaan field wajib seperti tanggal, nominal, akun/kategori, atau nomor referensi yang kosong atau tidak valid.",
         "Data kosong menunjukkan lemahnya validasi sistem input atau potensi penghilangan jejak transaksi.",
         "Lengkapi data yang hilang dari sumber transaksi asli sebelum kesimpulan akhir diambil."),
        ("A-14", "Menengah", "Observasi Nilai Negatif & Transaksi Retur/Koreksi",
         "Identifikasi nominal minus pada kolom pengeluaran, pendapatan, atau saldo yang seharusnya positif atau memiliki aturan tanda khusus.",
         "Nilai negatif bisa merupakan retur sah, reklasifikasi, atau koreksi/penyesuaian yang perlu otorisasi memadai.",
         "Pastikan setiap nilai negatif didukung nota retur, memo kredit, atau jurnal koreksi yang disetujui."),
        ("A-15", "Menengah", "Analisis Transaksi di Luar Hari Kerja (Weekend / Hari Libur)",
         "Penelusuran transaksi yang tercatat pada hari Sabtu, Minggu, atau hari libur resmi.",
         "Transaksi akhir pekan dapat mengindikasikan bypass kontrol harian, lemahnya pengawasan, atau jurnal penyesuaian tak wajar.",
         "Cek otorisasi, waktu posting sistem, dan alasan bisnis transaksi pada akhir pekan."),
        ("A-16", "Menengah", "Pencarian Kata Kunci Mencurigakan (Adjustment, Void, Koreksi)",
         "Pemindaian keterangan/uraian transaksi untuk kata kunci seperti adjustment, adj, void, koreksi, manual, atau write-off.",
         "Jurnal manual dan penyesuaian akhir periode memiliki risiko manipulasi atau koreksi kesalahan yang lebih tinggi.",
         "Minta alasan tertulis dan dokumen pendukung sah untuk setiap jurnal penyesuaian/pembatalan."),
        ("A-17", "Menengah", "Pengujian Overbudget (Realisasi Melebihi Anggaran/Pagu)",
         "Perbandingan realisasi aktual dengan pagu anggaran per pos/akun untuk menemukan pos yang melewati batas anggaran.",
         "Overbudget dapat menandakan pengendalian anggaran yang lemah, perubahan ruang lingkup tanpa revisi, atau pemborosan.",
         "Klarifikasi penyebab selisih anggaran dan cek persetujuan revisi pagu yang berwenang."),
        ("A-18", "Lanjutan", "Fraud Triangle & Red Flags Kecurangan Keuangan",
         "Tiga unsur pendorong fraud: Tekanan/Insentif (Pressure), Kesempatan (Opportunity), dan Rasionalisasi (Rationalization).",
         "Membantu auditor membaca pola perilaku dan celah sistem yang sering muncul sebelum kecurangan terungkap.",
         "Fokuskan pengujian tambahan pada area dengan tekanan target tinggi dan kontrol lemah."),
        ("A-19", "Lanjutan", "Pengujian Cut-Off (Pisah Batas Periode)",
         "Memastikan transaksi dan peristiwa dicatat pada periode akuntansi yang tepat sesuai penyerahan barang/jasa atau terjadinya kewajiban.",
         "Kesalahan cut-off menggeser pendapatan atau beban antarperiode sehingga laba dan posisi keuangan terdistorsi.",
         "Uji transaksi beberapa hari sebelum dan sesudah tanggal tutup buku terhadap bukti pengiriman/penerimaan."),
        ("A-20", "Lanjutan", "Prosedur Analitis Substantif & Ekspektasi Auditor",
         "Perbandingan angka keuangan dengan ekspektasi auditor berbasis tren historis, anggaran, data non-keuangan, atau rasio industri.",
         "Mengungkap fluktuasi tidak biasa atau hubungan yang tidak konsisten antarpos laporan keuangan.",
         "Tetapkan ambang selisih yang material dan selidiki varians yang menyimpang dari ekspektasi."),
        ("A-21", "Lanjutan", "Audit Kas & Rekonsiliasi Bank",
         "Pengujian saldo kas di tangan dan kas di bank melalui cash opname, konfirmasi bank, serta penelusuran item rekonsiliasi (deposit in transit, outstanding checks).",
         "Kas merupakan akun paling likuid dan paling rentan terhadap pencurian, lapping, atau window dressing.",
         "Telusuri pos rekonsiliasi lama/tidak wajar dan pastikan tidak ada cek beredar fiktif."),
        ("A-22", "Lanjutan", "Audit Piutang, Cadangan Kerugian & Penagihan",
         "Pengujian eksistensi piutang, ketepatan umur piutang (aging), dan kewajaran cadangan kerugian kredit ekspektasian.",
         "Piutang tak tertagih yang tidak dicadangkan atau penjualan fiktif dapat melebihsajikan aset dan laba.",
         "Lakukan konfirmasi saldo pelanggan, uji penerimaan kas setelah tanggal pelaporan, dan tinjau piutang macet."),
        ("A-23", "Lanjutan", "Audit Persediaan, Opname Fisik & Penilaian (NRV)",
         "Pengujian keberadaan, kuantitas, kepemilikan, dan penilaian persediaan pada nilai terendah antara biaya perolehan dan nilai realisasi neto.",
         "Barang usang/rusak atau kesalahan perhitungan fisik dapat mendistorsi HPP dan nilai persediaan.",
         "Amati stock opname, uji item bernilai tinggi/bergerak lambat, dan cek penurunan nilai persediaan."),
        ("A-24", "Lanjutan", "Audit Liabilitas & Pencarian Kewajiban yang Belum Tercatat",
         "Memastikan seluruh utang usaha, akrual, dan kewajiban kontinjensi telah diakui secara lengkap pada akhir periode.",
         "Kewajiban yang tidak tercatat membuat utang dan beban tampak lebih kecil sehingga laba terlihat lebih tinggi.",
         "Periksa pembayaran kas setelah akhir periode dan faktur tagihan yang belum dibukukan."),
        ("A-25", "Lanjutan", "Penyusunan Laporan Hasil Audit, Tindak Lanjut & Pemantauan",
         "Merangkum temuan secara objektif, menyusun rekomendasi praktis yang terukur, dan memantau komitmen perbaikan manajemen.",
         "Temuan audit baru memberi nilai tambah nyata jika diikuti perbaikan sistem dan pengendalian yang berkelanjutan.",
         "Susun rencana aksi dengan penanggung jawab dan tenggat waktu yang jelas untuk tiap temuan."),
    ]

    keuangan_25 = [
        ("K-01", "Dasar", "Kerangka Laporan Keuangan: Laba Rugi, Neraca & Arus Kas",
         "Laporan Laba Rugi mengukur kinerja periode; Neraca/Posisi Keuangan mengukur aset, liabilitas, dan ekuitas pada titik waktu; Arus Kas melacak aliran kas.",
         "Ketiganya saling terhubung—laba akuntansi dan arus kas memiliki perbedaan waktu dan non-kas sehingga harus dibaca bersama.",
         "Bandingkan tren laba dengan arus kas operasi sebelum menyimpulkan kualitas kinerja."),
        ("K-02", "Dasar", "Persamaan Dasar Akuntansi & Struktur Neraca",
         "Aset = Liabilitas + Ekuitas. Aset diklasifikasikan menjadi lancar dan tidak lancar; liabilitas menjadi jangka pendek dan jangka panjang.",
         "Menunjukkan sumber daya yang dikendalikan entitas dan bagaimana aset tersebut dibiayai (utang vs modal sendiri).",
         "Periksa keseimbangan neraca serta komposisi aset lancar vs liabilitas jangka pendek."),
        ("K-03", "Dasar", "Struktur Laporan Laba Rugi & Margin Utama",
         "Pendapatan dikurangi Beban Pokok/Operasional menghasilkan Laba Bruto, Laba Usaha/Operasional, hingga Laba Bersih.",
         "Membantu memisahkan efisiensi produksi/jasa inti, beban operasional, beban bunga/pajak, dan pos non-operasional.",
         "Pantau pergerakan margin kotor, margin operasi, dan margin bersih lintas periode."),
        ("K-04", "Dasar", "Analisis Horizontal (Tren Antarperiode)",
         "Membandingkan pos laporan keuangan dari satu periode ke periode lain dalam nominal maupun persentase perubahan.",
         "Mengidentifikasi percepatan pertumbuhan, kontraksi penjualan, atau lonjakan beban yang mulai terbentuk.",
         "Hitung pertumbuhan YoY/periode dan cek apakah pertumbuhan didorong volume/harga atau pos satu kali."),
        ("K-05", "Dasar", "Analisis Vertikal (Common-Size Statement)",
         "Menyatakan setiap pos sebagai persentase dari total acuan (misalnya % dari Total Aset pada neraca atau % dari Pendapatan pada laba rugi).",
         "Memudahkan perbandingan struktur biaya dan komposisi aset lintas ukuran entitas atau lintas waktu.",
         "Cari pos yang porsinya membesar secara tidak wajar dibanding total pendapatan atau total aset."),
        ("K-06", "Menengah", "Rasio Likuiditas: Current Ratio",
         "Current Ratio = Aset Lancar / Liabilitas Lancar. Mengukur kemampuan memenuhi kewajiban jangka pendek dengan aset lancar.",
         "Menilai bantalan likuiditas jangka pendek tanpa harus menjual aset tetap.",
         "Bandingkan dengan standar sektor serta perhatikan kualitas komponen aset lancar (kas vs piutang/persediaan)."),
        ("K-07", "Menengah", "Quick Ratio (Acid-Test Ratio) & Kualitas Likuiditas",
         "Quick Ratio = (Kas + Setara Kas + Piutang Lancar) / Liabilitas Lancar (mengeluarkan persediaan yang kurang likuid).",
         "Memberikan gambaran likuiditas yang lebih konservatif bila persediaan lambat berputar atau sulit dicairkan cepat.",
         "Waspadai current ratio yang tampak tinggi hanya karena penumpukan persediaan atau piutang lama."),
        ("K-08", "Menengah", "Modal Kerja Bersih (Net Working Capital) & Siklus Kas",
         "NWC = Aset Lancar - Liabilitas Lancar. Siklus konversi kas dipengaruhi perputaran piutang, persediaan, dan utang usaha.",
         "Menunjukkan kebutuhan kas operasional untuk memutar bisnis sehari-hari seiring pertumbuhan penjualan.",
         "Pertumbuhan pesat tanpa pengelolaan modal kerja dapat memicu tekanan kas meski laba tercatat positif."),
        ("K-09", "Menengah", "Rasio Profitabilitas: Net Profit Margin (NPM)",
         "NPM = (Laba Bersih / Pendapatan) × 100%. Mengukur berapa persen pendapatan yang berubah menjadi laba bersih.",
         "Mencerminkan kemampuan entitas menghasilkan laba akhir setelah seluruh beban, bunga, dan pajak.",
         "Evaluasi NPM bersama margin operasi agar dampak pajak atau pos non-operasional tidak menyesatkan."),
        ("K-10", "Menengah", "Return on Assets (ROA)",
         "ROA = (Laba Bersih / Total Aset) × 100%. Mengukur efisiensi penggunaan seluruh aset dalam menghasilkan laba.",
         "Menilai produktivitas basis aset entitas terlepas dari cara pendanaannya.",
         "Bandingkan ROA dengan biaya modal/bunga untuk menilai apakah aset menghasilkan imbal hasil memadai."),
        ("K-11", "Menengah", "Return on Equity (ROE)",
         "ROE = (Laba Bersih / Ekuitas) × 100%. Mengukur tingkat imbal hasil bagi pemilik modal/ekuitas.",
         "Menunjukkan seberapa efektif modal pemegang saham dikelola untuk menghasilkan laba.",
         "Cek apakah ROE tinggi berasal dari kinerja operasi yang kuat atau dari leverage utang yang tinggi."),
        ("K-12", "Menengah", "Rasio Solvabilitas: Debt to Equity Ratio (DER)",
         "DER = Total Liabilitas / Ekuitas. Mengukur proporsi pendanaan utang dibandingkan modal sendiri.",
         "Menggambarkan struktur permodalan dan ketahanan menghadapi tekanan keuangan atau kenaikan suku bunga.",
         "Nilai kewajaran DER sesuai karakteristik industri (utilitas/keuangan berbeda dengan perdagangan/jasa)."),
        ("K-13", "Menengah", "Debt to Asset Ratio & Kemampuan Bayar Bunga",
         "Debt to Asset = Total Liabilitas / Total Aset; Interest Coverage (EBIT / Beban Bunga) mengukur kemampuan menutup beban bunga.",
         "Membantu menilai risiko gagal bayar dan ruang gerak keuangan saat arus kas melemah.",
         "Perhatikan jadwal jatuh tempo utang dan sensitivitas laba terhadap perubahan beban bunga."),
        ("K-14", "Menengah", "Rasio Aktivitas & Efisiensi Perputaran Aset",
         "Asset Turnover = Pendapatan / Total Aset, ditambah perputaran piutang, persediaan, dan utang usaha.",
         "Menunjukkan seberapa cepat aset berputar menjadi penjualan dan kas.",
         "Perputaran yang melambat dapat menandakan kapasitas menganggur, stok menumpuk, atau penagihan lambat."),
        ("K-15", "Menengah", "Analisis DuPont (Dekomposisi ROE)",
         "ROE = Net Profit Margin × Total Asset Turnover × Equity Multiplier (Total Aset / Ekuitas).",
         "Mengurai sumber perubahan ROE menjadi profitabilitas, efisiensi aset, dan leverage keuangan.",
         "Gunakan dekomposisi ini untuk menentukan apakah perbaikan harus fokus ke margin, perputaran, atau struktur modal."),
        ("K-16", "Menengah", "Kualitas Laba (Earnings Quality) & Akrual",
         "Menilai sejauh mana laba didukung arus kas operasi dan bebas dari lonjakan akrual atau estimasi agresif.",
         "Laba akuntansi tinggi dengan arus kas operasi lemah dapat menjadi sinyal awal tekanan kualitas laba.",
         "Bandingkan laba bersih dengan arus kas dari aktivitas operasi secara konsisten."),
        ("K-17", "Menengah", "Analisis Anggaran vs Realisasi (Variance Analysis)",
         "Membandingkan anggaran/target dengan realisasi untuk menghitung varians nominal dan persentase pencapaian.",
         "Membantu manajemen mengevaluasi akurasi perencanaan, disiplin eksekusi, dan kebutuhan penyesuaian.",
         "Pisahkan varians akibat volume/aktivitas dari varians akibat harga/tarif atau efisiensi biaya."),
        ("K-18", "Lanjutan", "Keuangan Sektor Publik / Pemerintahan: Efektivitas & Efisiensi",
         "Efektivitas Pendapatan = (Realisasi Pendapatan / Anggaran Pendapatan) × 100%; Efisiensi Belanja = (Realisasi Belanja / Anggaran Belanja) × 100%.",
         "Mengukur ketepatan target penerimaan dan tingkat penyerapan/penghematan belanja publik.",
         "Interpretasikan efisiensi belanja bersama capaian output/outcome program agar tidak keliru menilai belanja rendah."),
        ("K-19", "Lanjutan", "Kemandirian Fiskal Daerah & Struktur Pendapatan Publik",
         "Rasio Kemandirian = (PAD / Pendapatan Transfer atau Total Pendapatan) × 100%.",
         "Menggambarkan kapasitas pendanaan sendiri daerah/entitas publik dibandingkan transfer eksternal.",
         "Dorong peningkatan PAD berkelanjutan melalui basis pajak/retribusi dan layanan publik yang sehat."),
        ("K-20", "Lanjutan", "Surplus/Defisit & Pembiayaan Netto",
         "Surplus/Defisit mencerminkan selisih pendapatan dan belanja periode berjalan; pembiayaan menutup defisit atau memanfaatkan surplus.",
         "Menunjukkan keseimbangan fiskal tahunan dan dampak terhadap saldo kas/SILPA.",
         "Pastikan penggunaan pembiayaan/SILPA terarah dan tidak menutupi defisit struktural berulang."),
        ("K-21", "Lanjutan", "Unit Economics, Break-Even & Struktur Biaya Tetap/Variabel",
         "Memisahkan biaya variabel (berubah seiring volume) dan biaya tetap untuk memahami titik impas dan sensitivitas laba.",
         "Membantu menentukan dampak perubahan volume penjualan terhadap laba dan kebutuhan kas.",
         "Hitung margin kontribusi sebelum menambah kapasitas atau diskon harga agresif."),
        ("K-22", "Lanjutan", "Valuasi & Kelipatan Keuangan (P/E, PBV, EV/EBITDA)",
         "Metrik penilaian berbasis laba, nilai buku ekuitas, atau nilai perusahaan terhadap EBITDA.",
         "Memberikan konteks kewajaran harga atau nilai relatif terhadap pertumbuhan dan risiko entitas.",
         "Gunakan beberapa metrik secara berdampingan dan sesuaikan dengan profil pertumbuhan serta risiko."),
        ("K-23", "Lanjutan", "Deteksi Dini Distres Keuangan & Risiko Kebangkrutan",
         "Pemantauan sinyal awal seperti rugi berulang, likuiditas mengetat, rasio utang membesar, dan arus kas operasi negatif.",
         "Memberi peringatan dini agar restrukturisasi atau efisiensi dapat dilakukan sebelum krisis likuiditas terjadi.",
         "Pantau indikator likuiditas dan arus kas bulanan, bukan hanya laporan tahunan."),
        ("K-24", "Lanjutan", "Catatan atas Laporan Keuangan (CALK) & Kebijakan Akuntansi",
         "Penjelasan naratif dan rincian kebijakan akuntansi, komitmen, kontinjensi, pihak berelasi, serta peristiwa setelah periode pelaporan.",
         "Angka di laporan utama sering kali tidak lengkap tanpa membaca kebijakan dan komitmen di CALK.",
         "Selalu baca CALK untuk memahami asumsi estimasi, jaminan/utang bersyarat, dan perubahan kebijakan."),
        ("K-25", "Lanjutan", "Menyusun Rekomendasi Keuangan yang Terukur",
         "Merumuskan langkah perbaikan yang spesifik, realistis, berprioritas, dan terukur berdasarkan temuan rasio dan tren.",
         "Menghubungkan analisis diagnostik menjadi keputusan manajemen yang konkret.",
         "Kaitkan setiap rekomendasi dengan target metrik, risiko yang dikurangi, dan dampak finansialnya."),
    ]

    aruskas_25 = [
        ("C-01", "Dasar", "Konsep Dasar Laporan Arus Kas (PSAK 207 / PSAK 2)",
         "Melaporkan penerimaan dan pengeluaran kas serta setara kas selama periode tertentu menurut aktivitas operasi, investasi, dan pendanaan.",
         "Menjelaskan bagaimana kas dihasilkan dan digunakan sehingga posisi kas awal berubah menjadi kas akhir.",
         "Bedakan aliran kas aktual dari pendapatan/beban berbasis akrual."),
        ("C-02", "Dasar", "Kas dan Setara Kas (Cash & Cash Equivalents)",
         "Kas di tangan/bank dan investasi jangka pendek yang sangat likuid, mudah dikonversi menjadi kas, dan berisiko kecil terhadap perubahan nilai (umumnya jatuh tempo ≤ 3 bulan).",
         "Menjadi dasar definisi kas yang direkonsiliasi dari awal hingga akhir periode pelaporan.",
         "Pastikan saldo kas awal dan akhir mengacu pada definisi kas dan setara kas yang konsisten."),
        ("C-03", "Dasar", "Aktivitas Operasi (Operating Cash Flows - CFO)",
         "Arus kas dari aktivitas penghasil pendapatan utama entitas dan aktivitas lain yang bukan investasi maupun pendanaan.",
         "Menunjukkan apakah bisnis inti mampu menghasilkan kas cukup untuk membiayai operasi tanpa mengandalkan pinjaman atau modal baru.",
         "Pantau konsistensi CFO positif sebagai indikator kesehatan operasional jangka panjang."),
        ("C-04", "Dasar", "Aktivitas Investasi (Investing Cash Flows - CFI)",
         "Arus kas dari perolehan dan pelepasan aset jangka panjang serta investasi lain yang tidak termasuk setara kas.",
         "Menggambarkan ekspansi kapasitas, peremajaan aset, atau pencairan investasi jangka panjang.",
         "Arus kas investasi negatif sering kali wajar jika mencerminkan belanja modal (CapEx) produktif."),
        ("C-05", "Dasar", "Aktivitas Pendanaan (Financing Cash Flows - CFF)",
         "Arus kas yang mengakibatkan perubahan dalam jumlah serta komposisi modal ekuitas dan pinjaman entitas.",
         "Menunjukkan bagaimana entitas mendanai pertumbuhan atau melunasi kewajiban pendanaan melalui utang, ekuitas, atau dividen.",
         "Perhatikan arus kas masuk pinjaman/modal vs arus kas keluar pembayaran pokok utang dan dividen."),
        ("C-06", "Menengah", "Metode Langsung (Direct Method) Arus Kas Operasi",
         "Mengungkapkan kelompok utama penerimaan kas bruto dan pembayaran kas bruto (misalnya kas diterima dari pelanggan, dibayar ke pemasok/karyawan/pajak).",
         "Memberikan informasi arus kas masuk dan keluar bruto yang bermanfaat untuk mengestimasi arus kas masa depan.",
         "Kelompokkan penerimaan dan pengeluaran kas bruto secara jelas dan konsisten."),
        ("C-07", "Menengah", "Metode Tidak Langsung (Indirect Method) Arus Kas Operasi",
         "Menyesuaikan laba/rugi bersih atas transaksi non-kas (penyusutan/amortisasi), pos akrual/deferal, serta perubahan aset dan liabilitas operasi.",
         "Menghubungkan laba bersih akuntansi dengan arus kas operasi serta menyoroti dampak modal kerja.",
         "Periksa perubahan piutang, persediaan, dan utang usaha yang memengaruhi konversi laba menjadi kas."),
        ("C-08", "Menengah", "Rekonsiliasi Saldo Kas Awal, Perubahan Neto & Kas Akhir",
         "Saldo Kas Akhir ≈ Saldo Kas Awal + Kenaikan/Penurunan Neto Kas (+ pengaruh perubahan kurs bila ada).",
         "Memastikan konsistensi matematis laporan arus kas dan mencegah selisih pencatatan yang tidak terjelaskan.",
         "Selidiki segera apabila saldo kas akhir laporan berbeda dari perhitungan rekonsiliasi saldo awal + perubahan neto."),
        ("C-09", "Menengah", "Pos Non-Kas & Transaksi Investasi/Pendanaan Non-Kas",
         "Transaksi seperti konversi utang menjadi saham, perolehan aset melalui sewa/kredit langsung, atau penyusutan yang tidak melibatkan kas periode berjalan.",
         "Mencegah kekeliruan mengira pos non-kas sebagai aliran kas aktual.",
         "Sajikan informasi transaksi non-kas secara terpisah sebagai pengungkapan pendukung."),
        ("C-10", "Menengah", "Klasifikasi Bunga, Dividen & Pajak Penghasilan",
         "Arus kas dari bunga dan dividen umumnya disajikan secara konsisten antarperiode; pajak penghasilan diklasifikasikan sebagai operasi kecuali dapat diidentifikasi khusus.",
         "Menjaga komparabilitas antarperiode dan ketepatan pengelompokan aktivitas.",
         "Pertahankan konsistensi klasifikasi bunga dan dividen dari tahun ke tahun."),
        ("C-11", "Menengah", "Free Cash Flow (FCF) & Makna Surplus/Defisit Kas Bebas",
         "FCF secara praktis diukur sebagai Arus Kas Operasi dikurangi Belanja Modal (CapEx) untuk aset tetap.",
         "Menunjukkan kas tersisa setelah membiayai operasi dan mempertahankan/mengembangkan aset tetap untuk melunasi utang, membayar dividen, atau menambah cadangan kas.",
         "Bedakan FCF negatif karena ekspansi CapEx terencana vs FCF negatif akibat CFO yang melemah."),
        ("C-12", "Menengah", "Cash Flow Margin (%)",
         "Cash Flow Margin = (Arus Kas Operasi / Pendapatan atau Penerimaan Pelanggan) × 100%.",
         "Mengukur seberapa efisien setiap rupiah pendapatan dikonversi menjadi kas operasi.",
         "Bandingkan dengan margin laba untuk melihat apakah penjualan cepat menjadi kas atau tertahan di piutang."),
        ("C-13", "Menengah", "Operating Cash Flow Ratio (x)",
         "Operating Cash Flow Ratio = Arus Kas Operasi / Liabilitas Lancar.",
         "Menilai kemampuan arus kas operasi menutupi kewajiban jangka pendek dalam satu periode.",
         "Rasio yang kuat menandakan pelunasan kewajiban jangka pendek ditopang kas operasional, bukan utang baru."),
        ("C-14", "Menengah", "Cash Flow to Debt Ratio (%)",
         "Cash Flow to Debt = (Arus Kas Operasi / Total Utang) × 100%.",
         "Mengukur kapasitas arus kas operasi untuk menutup keseluruhan kewajiban utang.",
         "Semakin tinggi rasio ini, semakin kuat kemampuan entitas melunasi utang dari kas operasional."),
        ("C-15", "Menengah", "Cash Flow Coverage / Debt Service Coverage (x)",
         "Mengukur kecukupan arus kas operasi terhadap pembayaran pokok pinjaman, bunga, atau kewajiban layanan utang.",
         "Menilai risiko tekanan kas saat jatuh tempo cicilan atau pembayaran bunga tiba.",
         "Perhatikan kecukupan kas operasi terutama saat jadwal pembayaran utang membesar."),
        ("C-16", "Menengah", "Analisis Tren Arus Kas Antarperiode (Komparatif)",
         "Membandingkan CFO, CFI, CFF, perubahan neto kas, dan saldo kas akhir lintas beberapa tahun/periode.",
         "Mengungkap pergeseran fase usaha, ketergantungan pendanaan, atau perubahan pola belanja modal.",
         "Lihat arah pergerakan tiap aktivitas secara terpisah, bukan hanya perubahan kas total."),
        ("C-17", "Menengah", "Analisis Tren Bulanan, Kas Masuk vs Kas Keluar & Burn Rate",
         "Menelusuri arus kas masuk, arus kas keluar, saldo berjalan bulanan, dan rata-rata pengeluaran kas operasional bulanan (burn rate).",
         "Membantu memantau volatilitas kas musiman dan memperkirakan ketahanan kas (cash runway).",
         "Waspadai pengeluaran bulanan yang konsisten melampaui kas masuk."),
        ("C-18", "Lanjutan", "Pola Siklus Hidup Perusahaan melalui Tanda CFO, CFI & CFF",
         "Kombinasi tanda (+/-) pada CFO, CFI, dan CFF sering mencerminkan tahapan pertumbuhan, ekspansi, kematangan, atau tekanan keuangan.",
         "Membantu menginterpretasikan apakah arus kas negatif pada satu aktivitas bersifat strategis atau bermasalah.",
         "Selalu gabungkan pola tanda arus kas dengan konteks bisnis dan posisi neraca."),
        ("C-19", "Lanjutan", "Risiko Ketergantungan pada Pendanaan Eksternal",
         "Terjadi ketika arus kas operasi negatif atau tidak mencukupi sehingga kas tetap terjaga berkat pinjaman baru atau setoran modal.",
         "Pendanaan eksternal bersifat terbatas dan membawa biaya bunga atau ekspektasi imbal hasil.",
         "Evaluasi rencana pemulihan CFO agar entitas tidak bergantung terus-menerus pada utang/modal baru."),
        ("C-20", "Lanjutan", "Modal Kerja & Dampak Piutang/Persediaan terhadap Kas",
         "Kenaikan piutang atau persediaan menyerap kas meskipun penjualan/laba tercatat naik; kenaikan utang usaha menahan kas keluar sementara.",
         "Menjelaskan mengapa laba bersih sering berbeda dari arus kas operasi.",
         "Periksa pertumbuhan piutang dan persediaan bersama pertumbuhan pendapatan."),
        ("C-21", "Lanjutan", "Deteksi Anomali Pengeluaran Kas (Metode IQR pada Mutasi)",
         "Mengidentifikasi transaksi pengeluaran kas yang berada di atas batas atas IQR (Q3 + 1,5×IQR) pada data mutasi/transaksi.",
         "Menyaring transaksi keluar bernilai ekstrem untuk ditinjau otorisasi, kewajaran, dan kelengkapan buktinya.",
         "Perlakukan hasil outlier sebagai daftar prioritas pemeriksaan, bukan vonis pelanggaran."),
        ("C-22", "Lanjutan", "Klasifikasi Transaksi Mutasi Bank & Keterbatasan Kata Kunci",
         "Mengelompokkan mutasi transaksi ke operasi, investasi, atau pendanaan berdasarkan keterangan/kategori.",
         "Klasifikasi otomatis berbasis kata kunci memiliki keterbatasan bila uraian transaksi singkat atau ambigu.",
         "Tinjau kembali klasifikasi transaksi bernilai besar atau beruraian umum sebelum finalisasi."),
        ("C-23", "Lanjutan", "Proksi & Estimasi Rasio saat Data Neraca/Laba Rugi Terbatas",
         "Menggunakan pos/proksi terkait di file saat akun neraca/laba rugi penuh tidak tersedia, sembari mencatat keterbatasannya.",
         "Memungkinkan pembacaan awal yang tetap transparan mengenai asumsi dan batasan data.",
         "Cantumkan secara jelas bahwa proksi/estimasi perlu diverifikasi dengan laporan keuangan lengkap."),
        ("C-24", "Lanjutan", "Manajemen Likuiditas, Buffer Kas & Perencanaan Arus Kas",
         "Menjaga cadangan kas memadai, menyelaraskan jadwal penerimaan dan pembayaran, serta menyusun proyeksi kas.",
         "Mencegah krisis likuiditas mendadak akibat keterlambatan penagihan atau pengeluaran tak terduga.",
         "Susun proyeksi arus kas jangka pendek dan tetapkan batas minimum kas operasional."),
        ("C-25", "Lanjutan", "Menyusun Ringkasan Eksekutif Arus Kas & Rekomendasi",
         "Menyajikan ringkasan yang memisahkan fakta angka, interpretasi, keterbatasan data, dan rekomendasi tindakan.",
         "Memudahkan pimpinan atau auditor mengambil keputusan berbasis bukti yang terukur.",
         "Pastikan setiap kesimpulan merujuk pada angka yang terverifikasi dan menyatakan asumsi dengan jujur."),
    ]

    return {
        "Audit Profesional": audit_25,
        "Analisis Keuangan": keuangan_25,
        "Analisis Arus Kas": aruskas_25,
    }


def _render_study_track(track_title, track_desc, materials):
    st.markdown(f"### {track_title}")
    st.caption(f"{track_desc} — Total **{len(materials)} materi** lengkap untuk pembelajaran.")
    c_f1, c_f2 = st.columns([2, 1])
    search_q = c_f1.text_input(
        f"Cari topik / kata kunci ({track_title})",
        placeholder="Contoh: sampling, likuiditas, free cash flow, rekonsiliasi...",
        key=f"study_search_{track_title}"
    )
    level_filter = c_f2.selectbox(
        "Filter tingkat materi",
        ["Semua Level", "Dasar", "Menengah", "Lanjutan"],
        key=f"study_level_{track_title}"
    )

    filtered = []
    for item in materials:
        code, lvl, title, konsep, fungsi, praktik = item
        if level_filter != "Semua Level" and lvl != level_filter:
            continue
        haystack = f"{code} {lvl} {title} {konsep} {fungsi} {praktik}".lower()
        if search_q.strip() and search_q.strip().lower() not in haystack:
            continue
        filtered.append(item)

    st.markdown(f"Menampilkan **{len(filtered)}** dari **{len(materials)}** materi.")
    if not filtered:
        st.info("Tidak ada materi yang cocok dengan filter pencarian.")
        return

    for code, lvl, title, konsep, fungsi, praktik in filtered:
        with st.expander(f"{code} [{lvl}] — {title}", expanded=False):
            st.markdown(f"**Konsep Inti:** {konsep}")
            st.markdown(f"**Kegunaan dalam Analisis:** {fungsi}")
            st.markdown(f"**Penerapan Praktis:** {praktik}")


def page_learning_center():
    finora_header("FINORA Study", "Pusat Pembelajaran Audit, Analisis Keuangan & Analisis Arus Kas")
    st.markdown(
        "Selamat datang di **FINORA Study**! Tersedia **75 materi pembelajaran** terstruktur "
        "(**25 materi per analisis**) untuk mendukung pembelajaran **Audit Profesional**, "
        "**Analisis Keuangan**, dan **Analisis Arus Kas**."
    )
    st.markdown("---")
    all_tracks = _get_finora_study_materials()
    tab_audit, tab_keuangan, tab_aruskas = st.tabs([
        "🔎 Audit Profesional (25 Materi)",
        "📈 Analisis Keuangan (25 Materi)",
        "💧 Analisis Arus Kas (25 Materi)"
    ])
    with tab_audit:
        _render_study_track(
            "Materi Audit Profesional",
            "Dasar audit, pengendalian intern, sampling, prosedur deteksi anomali, hingga pelaporan KKA",
            all_tracks["Audit Profesional"]
        )
    with tab_keuangan:
        _render_study_track(
            "Materi Analisis Keuangan",
            "Laporan keuangan, rasio likuiditas/profitabilitas/solvabilitas/aktivitas, sektor publik & evaluasi",
            all_tracks["Analisis Keuangan"]
        )
    with tab_aruskas:
        _render_study_track(
            "Materi Analisis Arus Kas",
            "PSAK 207/2, aktivitas operasi/investasi/pendanaan, rekonsiliasi, rasio arus kas, tren & risiko",
            all_tracks["Analisis Arus Kas"]
        )

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
            "FINORA Study", 
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
        elif nav == "FINORA Study": page_learning_center()
        elif nav == "Simpan": page_simpan()

if __name__ == "__main__":
    main()
