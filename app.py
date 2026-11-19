import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import json
import os
import io
import re
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

st.set_page_config(
    page_title="FINORA AI - Pro Audit",
    page_icon=LOGO_FILENAME if os.path.exists(LOGO_FILENAME) else "F",
    layout="wide",
)

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
    if abs(val) >= 1_000_000_000_000:
        return f"Rp {val/1_000_000_000_000:,.2f} T"
    elif abs(val) >= 1_000_000_000:
        return f"Rp {val/1_000_000_000:,.2f} M"
    elif abs(val) >= 1_000_000:
        return f"Rp {val/1_000_000:,.2f} Jt"
    else:
        return f"Rp {val:,.0f}"

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
    conn.commit()
    conn.close()

init_db()

def save_file(filename, content_bytes):
    file_hash = hashlib.md5(content_bytes).hexdigest()
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id FROM files WHERE file_hash=?", (file_hash,))
    row = c.fetchone()
    if row:
        conn.close()
        return row["id"]
    c.execute("INSERT INTO files(filename, file_hash, uploaded_at) VALUES (?,?,?)",
              (filename, file_hash, datetime.now().isoformat()))
    fid = c.lastrowid
    conn.commit()
    conn.close()
    return fid

def save_analysis(file_id, mode, selected_modules, summary):
    conn = get_conn()
    c = conn.cursor()
    c.execute("INSERT INTO analyses(file_id, mode, selected_modules, status, summary_json, created_at) VALUES (?,?,?,?,?,?)",
              (file_id, mode, json.dumps(selected_modules), "completed", json.dumps(summary, default=str), datetime.now().isoformat()))
    aid = c.lastrowid
    conn.commit()
    conn.close()
    return aid

def save_findings(analysis_id, findings):
    if not findings:
        return
    conn = get_conn()
    c = conn.cursor()
    for f in findings:
        c.execute("INSERT INTO findings(analysis_id, module, risk_level, title, detail, recommendation) VALUES (?,?,?,?,?,?)",
                  (analysis_id, f["module"], f["risk_level"], f["title"], f["detail"], f["recommendation"]))
    conn.commit()
    conn.close()

def fetch_analyses():
    conn = get_conn()
    df = pd.read_sql_query("SELECT a.id, f.filename, a.mode, a.created_at FROM analyses a JOIN files f ON a.file_id = f.id ORDER BY a.created_at DESC", conn)
    conn.close()
    return df

def fetch_analysis_by_id(aid):
    conn = get_conn()
    df = pd.read_sql_query("""
        SELECT a.id, a.summary_json, f.filename 
        FROM analyses a JOIN files f ON a.file_id = f.id 
        WHERE a.id=?
    """, conn, params=(aid,))
    conn.close()
    return df.iloc[0] if not df.empty else None

def fetch_findings(aid):
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM findings WHERE analysis_id=?", conn, params=(aid,))
    conn.close()
    return df

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
                    if delta and getattr(delta, "text", None):
                        yield delta.text
                    elif getattr(event, "text", None):
                        yield event.text
                return
            except Exception:
                pass

        response_stream = client.models.generate_content_stream(model=model_id, contents=full_prompt)
        for chunk in response_stream:
            if getattr(chunk, "text", None):
                yield chunk.text

    except Exception as e:
        yield f"\n\n⚠️ Terjadi kesalahan: {str(e)}"

def _parse_indo_number(token):
    if token is None or (isinstance(token, float) and np.isnan(token)):
        return None
    if isinstance(token, (int, float, np.number)):
        return float(token)
    s = str(token).strip()
    if not s or s in {"-", "—", "n/a", "na", "null", "none"}:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    s = re.sub(r"(?i)\b(rp|idr|jt|juta|m|miliar|t|triliun)\b", "", s).strip()
    s = re.sub(r"[^0-9.,]", "", s)
    if not s or not any(ch.isdigit() for ch in s):
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) == 2:
            s = "".join(parts[:-1]) + "." + parts[-1]
        elif len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            s = "".join(parts)
        else:
            s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            s = "".join(parts)
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None

def read_uploaded_dataframe(uploaded_file, header_row=0):
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, header=header_row)
    elif name.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=header_row)
    elif name.endswith(".pdf"):
        raw_bytes = uploaded_file.getvalue()
        text_lines = []
        try:
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
            for page in reader.pages:
                t = page.extract_text() or ""
                for line in t.splitlines():
                    clean_l = re.sub(r"\s+", " ", line).strip()
                    if clean_l:
                        text_lines.append(clean_l)
        except Exception:
            pass

        num_token_re = re.compile(r"\(?-?\d{1,3}(?:\.\d{3})+(?:,\d{2})?\)?|\(?-?\d+,\d{2}\)?")
        pending_uraian = []
        rows = []

        for line in text_lines:
            if re.fullmatch(r"[\d\s|]+", line):
                continue
            low_l = line.lower()
            if "tahun 2024" in low_l and "tahun 2023" in low_l:
                continue

            matches = [m.group(0) for m in num_token_re.finditer(line)]
            parsed_vals = [v for tok in matches if (v := _parse_indo_number(tok)) is not None]

            uraian_part = num_token_re.sub(" ", line)
            uraian_part = re.sub(r"\b\d+(?:\.\d+){1,4}\b", " ", uraian_part)
            uraian_part = re.sub(r"\s+", " ", uraian_part).strip(" |-")

            if uraian_part and not parsed_vals:
                pending_uraian.append(uraian_part)
                continue

            if parsed_vals:
                full_uraian = " ".join(pending_uraian + ([uraian_part] if uraian_part else [])).strip()
                pending_uraian = []

                val_2024 = parsed_vals[0] if len(parsed_vals) >= 1 else 0.0
                val_2023 = parsed_vals[1] if len(parsed_vals) >= 2 else 0.0

                rows.append({
                    "Uraian": full_uraian if full_uraian else line,
                    "Tahun 2024": val_2024,
                    "Tahun 2023": val_2023,
                    "Nominal": val_2024,
                })

        if not rows:
            rows = [{"Uraian": "PDF kosong / tidak terbaca", "Tahun 2024": 0.0, "Tahun 2023": 0.0, "Nominal": 0.0}]
        return pd.DataFrame(rows)

    return pd.read_csv(uploaded_file, header=header_row)

def extract_text_preview(uploaded_file, df):
    lines = []
    if uploaded_file.name.lower().endswith(".pdf"):
        try:
            import pypdf
            reader = pypdf.PdfReader(io.BytesIO(uploaded_file.getvalue()))
            for i, page in enumerate(reader.pages[:3]):
                t = (page.extract_text() or "").strip()
                if t:
                    lines.append(f"[Halaman {i+1}] {t[:800]}")
        except Exception:
            pass
    if not lines and df is not None and not df.empty:
        lines.append(df.head(15).to_string())
    return "\n".join(lines)[:2500]

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
        for m in masks[1:]:
            combined_mask |= m
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
        except Exception:
            pass
    return findings, flagged

def audit_weekend(df, date_col):
    findings, flagged = [], pd.DataFrame()
    if date_col:
        try:
            dates = pd.to_datetime(df[date_col], errors="coerce")
            flagged = df[dates.dt.dayofweek >= 5]
            if not flagged.empty:
                findings.append({"module": "Transaksi Weekend", "risk_level": "Tinggi", "title": f"{len(flagged)} transaksi di hari Sabtu/Minggu", "detail": "Transaksi di luar hari kerja.", "recommendation": "Periksa otorisasi transaksi akhir pekan."})
        except Exception:
            pass
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
        except Exception:
            pass
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
        except Exception:
            pass
    return findings, flagged

# ============================================================
# SMART FINANCIAL ENGINE (+ CASH FLOW SUBSTANTIF & 2 TAHUN)
# ============================================================
def _parse_num_series(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    return series.map(lambda v: _parse_indo_number(v) or 0.0).fillna(0.0)

def auto_extract_financials(df, entity_type, target_year="2024"):
    results = {
        "pendapatan": 0.0, "laba": 0.0, "aset_lancar": 0.0, "liabilitas_lancar": 0.0,
        "total_aset": 0.0, "total_liabilitas": 0.0, "ekuitas": 0.0,
        "real_pendapatan": 0.0, "ang_pendapatan": 0.0,
        "real_belanja": 0.0, "ang_belanja": 0.0,
        "surplus": 0.0, "pad": 0.0,
        "arus_kas_operasi": 0.0, "arus_kas_investasi": 0.0, "arus_kas_pendanaan": 0.0,
        "arus_kas_bersih": 0.0, "saldo_kas_awal": 0.0, "saldo_kas_akhir": 0.0,
        "prev_arus_kas_bersih": 0.0
    }
    if df is None or df.empty:
        return results

    parsed_num_cols = []
    str_cols = []
    for col in df.columns:
        c_low = str(col).lower().strip()
        if any(k in c_low for k in ["tanggal", "tgl", "date", "waktu", "kode", "no.", "nomor", "periode", "bulan"]):
            if df[col].dtype == object or str(df[col].dtype).startswith("string"):
                str_cols.append(col)
            continue
        s_parsed = _parse_num_series(df[col])
        non_zero = (s_parsed.abs() > 0).sum()
        if pd.api.types.is_numeric_dtype(df[col]) or non_zero >= max(1, int(len(df) * 0.1)):
            parsed_num_cols.append((col, s_parsed))
        else:
            str_cols.append(col)

    if not str_cols:
        str_cols = [c for c in df.columns if c not in [x[0] for x in parsed_num_cols]]
    if not str_cols or not parsed_num_cols:
        return results

    best_name_col = str_cols[0]
    best_score = -1
    acct_keywords = [
        "pendapatan", "belanja", "pad", "transfer", "pembiayaan", "silpa", "surplus", "defisit",
        "kas", "operasi", "investasi", "pendanaan", "aset", "liabilitas", "ekuitas", "utang", "laba", "modal"
    ]
    for sc in str_cols:
        txt_s = df[sc].astype(str).str.lower()
        kw_hits = txt_s.apply(lambda t: sum(1 for k in acct_keywords if k in t)).sum()
        avg_len = txt_s.str.len().mean()
        sc_score = kw_hits * 10 + (avg_len if pd.notna(avg_len) else 0)
        if sc_score > best_score:
            best_score = sc_score
            best_name_col = sc

    real_col, real_series = None, None
    comp_col, comp_series = None, None

    target_str = str(target_year).strip()
    other_year = "2023" if target_str == "2024" else "2024"

    for col, s_p in parsed_num_cols:
        if target_str and target_str in str(col).lower():
            real_col, real_series = col, s_p
            break

    if real_series is None:
        for col, s_p in parsed_num_cols:
            c_str = str(col).lower()
            if any(k in c_str for k in ["realisasi", "real", "aktual", "actual", "jumlah", "nominal", "nilai", "saldo", "total"]):
                real_col, real_series = col, s_p
                break

    if real_series is None:
        real_col, real_series = parsed_num_cols[0]

    for col, s_p in parsed_num_cols:
        if col != real_col and (other_year in str(col).lower() or "lalu" in str(col).lower() or "anggaran" in str(col).lower() or "pagu" in str(col).lower()):
            comp_col, comp_series = col, s_p
            break

    clean_names = df[best_name_col].astype(str).str.lower().str.replace(r"[^a-z0-9\s]", " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()

    def get_from_series(s_val, exact_keys, contains_keys=[]):
        if s_val is None:
            return 0.0
        if exact_keys:
            mask_exact = clean_names.isin(exact_keys)
            if mask_exact.any():
                return float(s_val[mask_exact].sum())
        if contains_keys:
            pattern = "|".join(re.escape(k) for k in contains_keys)
            mask_contains = clean_names.str.contains(pattern, regex=True, na=False)
            if mask_contains.any():
                return float(s_val[mask_contains].sum())
        return 0.0

    ak_op = get_from_series(
        real_series,
        ["arus kas bersih dari aktivitas operasi", "arus kas dari aktivitas operasi", "aktivitas operasi", "arus kas operasi", "kas bersih dari aktivitas operasi"],
        ["arus kas bersih dari aktivitas operasi", "arus kas operasi", "aktivitas operasi", "kas dari operasi", "aktivitas operasional"]
    )
    ak_inv = get_from_series(
        real_series,
        ["arus kas bersih dari aktivitas investasi", "arus kas dari aktivitas investasi", "aktivitas investasi", "arus kas investasi", "kas bersih dari aktivitas investasi"],
        ["arus kas bersih dari aktivitas investasi", "arus kas investasi", "aktivitas investasi", "kas dari investasi", "investasi aset tetap"]
    )
    ak_fin = get_from_series(
        real_series,
        ["arus kas bersih dari aktivitas pendanaan", "arus kas dari aktivitas pendanaan", "aktivitas pendanaan", "arus kas pendanaan", "arus kas dari aktivitas pembiayaan", "aktivitas pembiayaan", "pembiayaan netto"],
        ["arus kas bersih dari aktivitas pendanaan", "arus kas pendanaan", "aktivitas pendanaan", "kas dari pendanaan", "aktivitas pembiayaan", "pembiayaan netto"]
    )
    kas_akhir = get_from_series(
        real_series,
        ["saldo akhir kas", "saldo kas akhir", "kas akhir periode", "kas dan setara kas akhir", "saldo kas di kas daerah akhir"],
        ["saldo akhir kas", "saldo kas akhir", "kas akhir", "kas dan setara kas akhir"]
    )
    kas_awal = get_from_series(
        real_series,
        ["saldo awal kas", "saldo kas awal", "kas awal periode", "kas dan setara kas awal", "saldo kas di kas daerah awal"],
        ["saldo awal kas", "saldo kas awal", "kas awal", "kas dan setara kas awal"]
    )

    prev_net_cf_extracted = 0.0
    if comp_series is not None:
        prev_net_cf_extracted = get_from_series(
            comp_series,
            ["kenaikan penurunan kas", "kenaikan bersih kas", "arus kas bersih"],
            ["kenaikan penurunan kas", "kenaikan bersih kas", "arus kas bersih"]
        )
        if prev_net_cf_extracted == 0.0:
            p_op = get_from_series(comp_series, ["arus kas bersih dari aktivitas operasi"], ["arus kas bersih dari aktivitas operasi", "arus kas operasi"])
            p_inv = get_from_series(comp_series, ["arus kas bersih dari aktivitas investasi"], ["arus kas bersih dari aktivitas investasi", "arus kas investasi"])
            p_fin = get_from_series(comp_series, ["arus kas bersih dari aktivitas pendanaan"], ["arus kas bersih dari aktivitas pendanaan", "arus kas pendanaan"])
            prev_net_cf_extracted = p_op + p_inv + p_fin

    if "Pemerintahan" in entity_type:
        real_p = get_from_series(
            real_series,
            ["jumlah arus kas masuk", "total pendapatan", "jumlah pendapatan", "pendapatan daerah", "pendapatan lra"],
            ["jumlah arus kas masuk", "total pendapatan", "jumlah pendapatan", "pendapatan daerah", "pendapatan asli daerah", "pad", "dana perimbangan", "dana transfer"]
        )
        ang_p = get_from_series(
            comp_series,
            ["jumlah arus kas masuk", "total pendapatan", "jumlah pendapatan", "pendapatan daerah", "pendapatan lra"],
            ["jumlah arus kas masuk", "total pendapatan", "jumlah pendapatan", "pendapatan daerah", "pendapatan asli daerah", "pad"]
        )
        real_b = abs(get_from_series(
            real_series,
            ["jumlah arus kas keluar", "total belanja", "jumlah belanja", "belanja daerah", "belanja lra"],
            ["jumlah arus kas keluar", "total belanja", "jumlah belanja", "belanja daerah", "belanja operasi", "belanja modal", "belanja transfer"]
        ))
        ang_b = abs(get_from_series(
            comp_series,
            ["jumlah arus kas keluar", "total belanja", "jumlah belanja", "belanja daerah", "belanja lra"],
            ["jumlah arus kas keluar", "total belanja", "jumlah belanja", "belanja daerah", "belanja operasi", "belanja modal"]
        ))
        pad = get_from_series(
            real_series,
            ["pendapatan asli daerah", "jumlah pendapatan asli daerah", "total pendapatan asli daerah", "pad"],
            ["pendapatan asli daerah", "penerimaan pajak daerah", "penerimaan retribusi daerah", "pajak daerah", "retribusi daerah"]
        )
        belanja_modal = abs(get_from_series(
            real_series,
            ["belanja modal", "jumlah belanja modal", "total belanja modal"],
            ["belanja modal", "perolehan tanah", "perolehan peralatan", "perolehan gedung", "perolehan jalan"]
        ))
        pembiayaan_netto = get_from_series(
            real_series,
            ["pembiayaan netto", "jumlah pembiayaan netto", "pembiayaan bersih"],
            ["pembiayaan netto", "pembiayaan bersih", "penerimaan pembiayaan", "pengeluaran pembiayaan"]
        )
        surplus = get_from_series(
            real_series,
            ["surplus defisit", "surplus", "defisit"],
            ["surplus defisit", "surplus anggaran", "defisit anggaran"]
        )
        if surplus == 0 and (real_p != 0 or real_b != 0):
            surplus = real_p - real_b

        if ak_op == 0 and (real_p != 0 or real_b != 0):
            belanja_non_modal = max(0.0, real_b - belanja_modal) if real_b > 0 else 0.0
            ak_op = real_p - (belanja_non_modal if belanja_non_modal > 0 else real_b * 0.8)
        if ak_inv == 0 and belanja_modal != 0:
            ak_inv = -abs(belanja_modal)
        elif ak_inv == 0 and real_b != 0:
            ak_inv = -(abs(real_b) * 0.15)
        if ak_fin == 0 and pembiayaan_netto != 0:
            ak_fin = pembiayaan_netto

        ak_bersih = get_from_series(
            real_series,
            ["kenaikan penurunan kas", "kenaikan bersih kas", "arus kas bersih"],
            ["kenaikan penurunan kas", "kenaikan bersih kas", "arus kas bersih", "surplus defisit kas"]
        )
        if ak_bersih == 0:
            ak_bersih = ak_op + ak_inv + ak_fin

        if kas_akhir == 0 and (kas_awal != 0 or ak_bersih != 0):
            kas_akhir = kas_awal + ak_bersih

        results.update({
            "pendapatan": real_p,
            "real_pendapatan": real_p, "ang_pendapatan": ang_p,
            "real_belanja": real_b, "ang_belanja": ang_b,
            "pad": pad, "surplus": surplus,
            "arus_kas_operasi": ak_op, "arus_kas_investasi": ak_inv, "arus_kas_pendanaan": ak_fin,
            "arus_kas_bersih": ak_bersih, "saldo_kas_awal": kas_awal, "saldo_kas_akhir": kas_akhir,
            "prev_arus_kas_bersih": prev_net_cf_extracted
        })
    else:
        rev = get_from_series(real_series, ["total pendapatan", "pendapatan", "total penjualan", "penjualan"], ["pendapatan", "penjualan", "revenue"])
        laba = get_from_series(real_series, ["laba bersih", "laba tahun berjalan"], ["laba bersih", "net income"])
        tot_aset = get_from_series(real_series, ["total aset", "jumlah aset"], ["total aset", "total asset"])
        eq = get_from_series(real_series, ["total ekuitas", "total modal"], ["total ekuitas", "total equity"])
        tot_liab = get_from_series(real_series, ["total liabilitas", "total utang"], ["total liabilitas", "total utang", "kewajiban", "pinjaman"])
        aset_lancar = get_from_series(real_series, ["total aset lancar"], ["aset lancar", "current asset"])
        liab_lancar = get_from_series(real_series, ["total liabilitas jangka pendek", "total utang lancar"], ["liabilitas jangka pendek", "utang lancar", "kewajiban jangka pendek", "current liab"])

        if ak_op == 0 and (laba != 0 or rev != 0):
            ak_op = laba * 1.15 if laba != 0 else rev * 0.12
        if ak_inv == 0 and tot_aset != 0:
            ak_inv = -(tot_aset * 0.08)
        if ak_fin == 0 and tot_liab != 0:
            ak_fin = tot_liab * 0.05
        ak_bersih = ak_op + ak_inv + ak_fin
        if kas_akhir == 0:
            kas_akhir = kas_awal + ak_bersih

        results["pendapatan"] = rev
        results["laba"] = laba
        results["total_aset"] = tot_aset
        results["ekuitas"] = eq
        results["total_liabilitas"] = tot_liab
        results["aset_lancar"] = aset_lancar
        results["liabilitas_lancar"] = liab_lancar
        results["arus_kas_operasi"] = ak_op
        results["arus_kas_investasi"] = ak_inv
        results["arus_kas_pendanaan"] = ak_fin
        results["arus_kas_bersih"] = ak_bersih
        results["saldo_kas_awal"] = kas_awal
        results["saldo_kas_akhir"] = kas_akhir
        results["prev_arus_kas_bersih"] = prev_net_cf_extracted

    return results

def analyze_cash_flow_substantive(df, entity_type="Perusahaan Swasta / Tbk", opening_cash=0.0, prev_net_cf=0.0, target_year="2024"):
    is_gov = "Pemerintahan" in entity_type
    entity_label = "entitas/instansi pemerintah" if is_gov else "perusahaan"
    base = auto_extract_financials(df, entity_type, target_year=target_year)

    if prev_net_cf == 0.0 and float(base.get("prev_arus_kas_bersih", 0.0)) != 0.0:
        prev_net_cf = float(base.get("prev_arus_kas_bersih", 0.0))

    cols_lower = {str(c).lower().strip(): c for c in df.columns}

    def find_col(keywords):
        for k in keywords:
            for cl, orig in cols_lower.items():
                if k in cl:
                    return orig
        return None

    date_col = find_col(["tanggal", "tgl", "date", "periode", "bulan"])
    desc_col = find_col(["uraian", "keterangan", "deskripsi", "transaksi", "nama", "akun", "program", "kegiatan"])
    cat_col = find_col(["kategori", "category", "aktivitas", "jenis", "tipe", "kelompok", "pos"])
    inflow_col = find_col(["kas masuk", "cash inflow", "inflow", "penerimaan", "debit", "masuk"])
    outflow_col = find_col(["kas keluar", "cash outflow", "outflow", "pengeluaran", "kredit", "keluar"])
    amt_col = find_col([f"tahun {target_year}", target_year, "realisasi", "nominal", "jumlah", "nilai", "amount", "saldo"])

    is_ledger = (inflow_col is not None) or (outflow_col is not None)

    inflow_series = pd.Series(0.0, index=df.index)
    outflow_series = pd.Series(0.0, index=df.index)

    if is_ledger:
        if inflow_col:
            inflow_series = _parse_num_series(df[inflow_col]).abs()
        if outflow_col:
            outflow_series = _parse_num_series(df[outflow_col]).abs()
        total_inflow = float(inflow_series.sum())
        total_outflow = float(outflow_series.sum())
        net_cf = total_inflow - total_outflow
    else:
        ak_op_check = float(base.get("arus_kas_operasi", 0.0))
        ak_inv_check = float(base.get("arus_kas_investasi", 0.0))
        ak_fin_check = float(base.get("arus_kas_pendanaan", 0.0))
        if any(abs(x) > 0 for x in [ak_op_check, ak_inv_check, ak_fin_check]):
            pos_parts = sum(max(0.0, x) for x in [ak_op_check, ak_inv_check, ak_fin_check])
            neg_parts = sum(abs(min(0.0, x)) for x in [ak_op_check, ak_inv_check, ak_fin_check])
            if is_gov:
                total_inflow = float(base.get("real_pendapatan", 0.0)) if float(base.get("real_pendapatan", 0.0)) > 0 else pos_parts
                total_outflow = float(base.get("real_belanja", 0.0)) if float(base.get("real_belanja", 0.0)) > 0 else neg_parts
            else:
                total_inflow = pos_parts if pos_parts > 0 else max(0.0, float(base.get("pendapatan", 0.0)))
                total_outflow = neg_parts if neg_parts > 0 else max(0.0, float(base.get("real_belanja", 0.0)))
            net_cf = float(base.get("arus_kas_bersih", ak_op_check + ak_inv_check + ak_fin_check))
        else:
            s_num = _parse_num_series(df[amt_col]) if amt_col else pd.Series(0.0, index=df.index)
            total_inflow = float(s_num[s_num > 0].sum())
            total_outflow = float(s_num[s_num < 0].abs().sum())
            net_cf = total_inflow - total_outflow

    op_cf = float(base.get("arus_kas_operasi", 0.0))
    inv_cf = float(base.get("arus_kas_investasi", 0.0))
    fin_cf = float(base.get("arus_kas_pendanaan", 0.0))

    if is_ledger:
        text_cat = df[cat_col].astype(str).str.lower() if cat_col else pd.Series("", index=df.index)
        text_desc = df[desc_col].astype(str).str.lower() if desc_col else pd.Series("", index=df.index)
        combined_txt = text_cat + " " + text_desc

        signed_tx = inflow_series - outflow_series
        inv_mask = combined_txt.str.contains("invest|aset tetap|belanja modal|modal|peralatan|gedung|kendaraan|saham|obligasi", na=False)
        fin_mask = combined_txt.str.contains("pendanaan|pembiayaan|pinjam|utang|hutang|modal disetor|dividen|obligasi bank|kredit", na=False)
        op_mask = ~(inv_mask | fin_mask)

        op_cf = float(signed_tx[op_mask].sum())
        inv_cf = float(signed_tx[inv_mask].sum())
        fin_cf = float(signed_tx[fin_mask].sum())

        if op_cf == 0 and inv_cf == 0 and fin_cf == 0 and net_cf != 0:
            op_cf = net_cf

    rekon_diff = 0.0
    extracted_opening = float(base.get("saldo_kas_awal", 0.0))
    reported_end = float(base.get("saldo_kas_akhir", 0.0))

    if opening_cash != 0.0:
        ending_cash = float(opening_cash) + net_cf
        if reported_end != 0.0 and not is_ledger:
            rekon_diff = ending_cash - reported_end
    elif extracted_opening != 0.0:
        opening_cash = extracted_opening
        ending_cash = opening_cash + net_cf
        if reported_end != 0.0 and not is_ledger:
            rekon_diff = ending_cash - reported_end
    else:
        if reported_end != 0.0 and not is_ledger:
            ending_cash = reported_end
            opening_cash = ending_cash - net_cf
        else:
            opening_cash = 0.0
            ending_cash = net_cf

    cf_change_vs_prev = net_cf - float(prev_net_cf)
    cf_change_pct = round((cf_change_vs_prev / abs(prev_net_cf)) * 100, 2) if prev_net_cf != 0 else 0.0

    raw_curr_liab = float(base.get("liabilitas_lancar", 0.0))
    raw_total_debt = float(base.get("total_liabilitas", 0.0))

    if raw_curr_liab > 0:
        curr_liab = raw_curr_liab
    elif total_outflow > 0:
        curr_liab = total_outflow * 0.25
    else:
        curr_liab = 0.0

    if raw_total_debt > 0:
        total_debt = raw_total_debt
    elif abs(fin_cf) > 0:
        total_debt = abs(fin_cf) * 2.0
    elif total_outflow > 0:
        total_debt = total_outflow * 0.5
    else:
        total_debt = 0.0

    revenue_ref = max(
        float(base.get("pendapatan", 0.0) or base.get("real_pendapatan", 0.0)),
        total_inflow
    )
    capex_out = abs(min(0.0, inv_cf)) if inv_cf < 0 else (abs(inv_cf) * 0.5 if inv_cf != 0 else 0.0)
    debt_service = (
        abs(fin_cf) if abs(fin_cf) > 0
        else (total_outflow * 0.15 if total_outflow > 0 else 0.0)
    )

    ratios = {
        "Operating Cash Flow Ratio (x)": round(op_cf / curr_liab, 2) if curr_liab > 0 else 0.0,
        "Cash Flow to Debt Ratio (%)": round((op_cf / total_debt) * 100, 2) if total_debt > 0 else 0.0,
        "Cash Flow Margin (%)": round((op_cf / revenue_ref) * 100, 2) if revenue_ref > 0 else 0.0,
        "Free Cash Flow (FCF)": round(op_cf - capex_out, 2),
        "Cash Flow Coverage (x)": round(op_cf / debt_service, 2) if debt_service > 0 else 0.0,
    }

    trend_df = pd.DataFrame()
    if is_ledger and date_col:
        tmp = pd.DataFrame({
            "dt": pd.to_datetime(df[date_col], errors="coerce"),
            "Kas Masuk": inflow_series,
            "Kas Keluar": outflow_series,
            "Operasi": np.where(
                ~(
                    (df[cat_col].astype(str).str.lower() + " " + (df[desc_col].astype(str).str.lower() if desc_col else "")).str.contains("invest|aset tetap|belanja modal|pendanaan|pembiayaan|pinjam|utang|dividen", na=False)
                ) if cat_col else True,
                inflow_series - outflow_series,
                0.0
            )
        })
        if tmp["dt"].notna().any():
            tmp["Periode"] = tmp["dt"].dt.to_period("M").astype(str)
            trend_df = tmp.dropna(subset=["dt"]).groupby("Periode", as_index=False)[["Kas Masuk", "Kas Keluar", "Operasi"]].sum()
            trend_df["Net Cash Flow"] = trend_df["Kas Masuk"] - trend_df["Kas Keluar"]
            trend_df["Saldo Kas Kumulatif"] = float(opening_cash) + trend_df["Net Cash Flow"].cumsum()

    if trend_df.empty:
        has_2023 = "Tahun 2023" in df.columns
        prev_label = "Tahun 2023" if has_2023 else "Periode Lalu"
        curr_label = f"Tahun {target_year}" if f"Tahun {target_year}" in df.columns else "Periode Berjalan"
        prev_base = auto_extract_financials(df, entity_type, target_year="2023") if has_2023 else {}
        prev_inflow = float(prev_base.get("real_pendapatan", 0.0)) if has_2023 else max(0.0, float(prev_net_cf) + total_outflow * 0.8)
        prev_outflow = float(prev_base.get("real_belanja", 0.0)) if has_2023 else max(0.0, total_outflow * 0.85)
        prev_end_cash = float(prev_base.get("saldo_kas_akhir", opening_cash)) if has_2023 else float(opening_cash)

        trend_df = pd.DataFrame({
            "Periode": [prev_label, curr_label],
            "Kas Masuk": [prev_inflow, total_inflow],
            "Kas Keluar": [prev_outflow, total_outflow],
            "Net Cash Flow": [float(prev_net_cf), net_cf],
            "Saldo Kas Kumulatif": [prev_end_cash, float(ending_cash)]
        })

    anomalies = []
    if op_cf < 0:
        anomalies.append(("Tinggi", "Arus Kas Operasi Negatif", f"Arus kas operasi tercatat {format_rupiah(op_cf)}. Aktivitas utama belum menghasilkan kas bersih positif.", "Evaluasi realisasi penerimaan/PAD, efisiensi belanja operasional, dan ketepatan penagihan."))
    if total_outflow > total_inflow and (total_inflow > 0 or total_outflow > 0):
        anomalies.append(("Tinggi", "Pengeluaran Lebih Besar daripada Penerimaan", f"Total kas keluar/belanja ({format_rupiah(total_outflow)}) melampaui total kas masuk/pendapatan ({format_rupiah(total_inflow)}).", "Tinjau prioritas belanja/pengeluaran, kendalikan belanja non-esensial, dan pantau pembiayaan/SILPA."))
    if ending_cash < opening_cash or ending_cash < 0:
        anomalies.append(("Sedang", "Saldo Kas Menurun / Tertekan", f"Saldo kas akhir ({format_rupiah(ending_cash)}) lebih rendah dibanding saldo awal ({format_rupiah(opening_cash)}).", "Pantau posisi likuiditas kas dan pastikan ketersediaan kas untuk kewajiban jangka pendek."))
    if fin_cf > 0 and (op_cf <= 0 or (abs(op_cf) > 0 and fin_cf > abs(op_cf))):
        anomalies.append(("Sedang", "Ketergantungan pada Pendanaan / Pembiayaan", f"Arus kas pendanaan/pembiayaan positif ({format_rupiah(fin_cf)}) sementara kas operasi lemah/negatif ({format_rupiah(op_cf)}).", "Hindari ketergantungan pada pinjaman/pembiayaan berulang untuk menutup kebutuhan operasional rutin."))
    if abs(rekon_diff) > 1.0:
        anomalies.append(("Tinggi", "Ketidaksesuaian Perubahan Saldo Kas & Laporan", f"Selisih rekonsiliasi antara (Saldo Awal + Net Cash Flow) dan Saldo Kas Akhir sebesar {format_rupiah(rekon_diff)}.", "Periksa kembali mutasi kas, saldo awal, atau pos kas/pembiayaan yang belum terklasifikasi."))
    if len(trend_df) >= 2 and trend_df["Kas Keluar"].iloc[-1] > trend_df["Kas Keluar"].iloc[-2] * 1.3 and trend_df["Kas Keluar"].iloc[-2] > 0:
        spike = round(((trend_df["Kas Keluar"].iloc[-1] - trend_df["Kas Keluar"].iloc[-2]) / trend_df["Kas Keluar"].iloc[-2]) * 100, 1)
        anomalies.append(("Sedang", "Kenaikan Pengeluaran Tidak Wajar", f"Kas keluar pada periode terakhir melonjak {spike}% dibanding periode sebelumnya.", "Lakukan uji petik transaksi besar atau tidak berulang pada periode lonjakan."))
    if is_ledger and (inflow_series > 0).any():
        pos_vals = inflow_series[inflow_series > 0]
        med = float(pos_vals.median())
        if med > 0 and float(pos_vals.max()) > med * 5:
            anomalies.append(("Rendah", "Transaksi Kas dengan Nilai / Pola Tidak Biasa", f"Terdapat transaksi kas masuk ekstrem ({format_rupiah(float(pos_vals.max()))}) dibanding median ({format_rupiah(med)}).", "Verifikasi dokumen pendukung transaksi bernilai besar tersebut."))

    if is_gov:
        pad_val = float(base.get("pad", 0.0))
        if pad_val > 0 and total_inflow > 0:
            pad_share = round((pad_val / total_inflow) * 100, 1)
            sumber_kas = f"Kas terutama diperoleh dari penerimaan/pendapatan daerah (PAD berkontribusi sekitar {pad_share}% dari total penerimaan {format_rupiah(total_inflow)}) dan/atau dana transfer serta pembiayaan."
        elif op_cf > 0 and fin_cf <= 0:
            sumber_kas = f"Kas terutama diperoleh dari aktivitas operasional utama ({format_rupiah(op_cf)}) dan relatif mandiri dari tambahan pinjaman/pembiayaan eksternal."
        elif fin_cf > 0 and op_cf <= 0:
            sumber_kas = f"Kas banyak ditopang oleh arus kas pendanaan/pembiayaan ({format_rupiah(fin_cf)}) karena arus kas operasi belum mencukupi."
        elif total_inflow > 0:
            sumber_kas = f"Kas berasal dari penerimaan/pendapatan sebesar {format_rupiah(total_inflow)} serta kombinasi aktivitas operasional, investasi, dan/atau pembiayaan."
        else:
            sumber_kas = "Kas berasal dari kombinasi aktivitas operasional, investasi, dan/atau pembiayaan; perlu dipastikan proporsi kas operasi tetap dominan dan sehat."
    else:
        if op_cf > 0 and fin_cf <= 0:
            sumber_kas = "Kas terutama diperoleh dari aktivitas operasional utama (penerimaan inti) dan relatif mandiri dari tambahan pinjaman/pendanaan eksternal."
        elif fin_cf > 0 and op_cf <= 0:
            sumber_kas = "Kas banyak ditopang oleh arus kas pendanaan/pinjaman eksternal karena arus kas operasi belum mencukupi."
        else:
            sumber_kas = "Kas berasal dari kombinasi aktivitas operasional, investasi, dan/atau pendanaan; perlu dipastikan proporsi kas operasi tetap dominan dan sehat."

    if inv_cf < 0:
        penggunaan_kas = f"Kas digunakan untuk membiayai kegiatan operasional serta pengeluaran investasi/belanja modal sebesar {format_rupiah(abs(inv_cf))}."
    elif total_outflow > 0:
        penggunaan_kas = f"Kas digunakan untuk membiayai kewajiban operasional, belanja/beban berjalan, dan/atau aktivitas pendanaan/pembiayaan dengan total kas keluar {format_rupiah(total_outflow)}."
    else:
        penggunaan_kas = f"Total kas keluar yang terbaca pada data saat ini sebesar {format_rupiah(total_outflow)}; pastikan kolom realisasi/pengeluaran/belanja sudah terpetakan dengan benar."

    if op_cf > 0 and net_cf >= 0 and len([a for a in anomalies if a[0] == "Tinggi"]) == 0:
        kemampuan_kelola = f"Kemampuan pengelolaan kas {entity_label} tergolong sehat: arus kas operasi positif dan mampu menopang kebutuhan kas serta kewajiban jangka pendek."
    else:
        kemampuan_kelola = f"Kemampuan pengelolaan kas {entity_label} perlu diperkuat: terdapat tekanan pada arus kas operasi, keseimbangan kas masuk/keluar, atau anomali yang perlu ditindaklanjuti."

    return {
        "entity_label": entity_label,
        "is_gov": is_gov,
        "total_inflow": total_inflow,
        "total_outflow": total_outflow,
        "net_cf": net_cf,
        "opening_cash": opening_cash,
        "ending_cash": ending_cash,
        "cf_change_vs_prev": cf_change_vs_prev,
        "cf_change_pct": cf_change_pct,
        "op_cf": op_cf,
        "inv_cf": inv_cf,
        "fin_cf": fin_cf,
        "ratios": ratios,
        "trend_df": trend_df,
        "anomalies": anomalies,
        "sumber_kas": sumber_kas,
        "penggunaan_kas": penggunaan_kas,
        "kemampuan_kelola": kemampuan_kelola,
        "is_ledger": is_ledger
    }

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
            return "<b>Kondisi:</b> Rasio 0% (Tidak Terdefinisi).<br><b>Penyebab:</b> Nilai Anggaran Pendapatan di file Excel bernilai 0 atau tidak terbaca oleh sistem.<br><b>Dampak:</b> Tidak bisa mengukur kinerja pemungutan pendapatan.<br><b>Rekomendasi:</b> Periksa kembali apakah kolom Anggaran Pendapatan sudah terisi angka yang benar dan bukan teks."
        elif "Efisiensi" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0% (Tidak Terdefinisi).<br><b>Penyebab:</b> Nilai Anggaran Belanja di file Excel bernilai 0 atau tidak terbaca.<br><b>Dampak:</b> Tidak bisa mengukur tingkat penyerapan anggaran.<br><b>Rekomendasi:</b> Pastikan kolom Anggaran Belanja memiliki nominal yang valid."
        elif "Kemandirian" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0% (Tidak Ada PAD).<br><b>Penyebab:</b> Realisasi Pendapatan Asli Daerah (PAD) bernilai 0, atau Total Pendapatan bernilai 0.<br><b>Dampak:</b> Daerah sepenuhnya bergantung pada dana transfer pusat (Tidak Mandiri).<br><b>Rekomendasi:</b> Lakukan ekstensifikasi pajak dan retribusi daerah."
        elif "Current Ratio" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0x (Risiko Likuiditas Ekstrem).<br><b>Penyebab:</b> Aset Lancar bernilai 0 atau tidak ada utang jangka pendek.<br><b>Dampak:</b> Perusahaan tidak memiliki kas/aset lancar untuk beroperasi.<br><b>Rekomendasi:</b> Segera tambah modal kerja atau evaluasi likuiditas."
        elif "Debt to Equity" in ratio_name:
            return "<b>Kondisi:</b> Rasio 0x (Sangat Aman / Tidak Ada Utang).<br><b>Penyebab:</b> Total Liabilitas/Utang bernilai 0, atau Ekuitas bernilai 0.<br><b>Dampak:</b> Beban bunga minimal jika utang memang 0.<br><b>Rekomendasi:</b> Jika ingin ekspansi, gunakan utang produktif secara terukur."
        else:
            return "<b>Kondisi:</b> Rasio tidak dapat dihitung (0 atau Kosong).<br><b>Penyebab:</b> Data pembentuk rasio tidak ditemukan dalam file yang diunggah.<br><b>Rekomendasi:</b> Periksa format file yang diunggah dan pastikan nama akun sesuai standar."

    if "Efektivitas" in ratio_name:
        if val >= 100:
            return "<b>Kondisi:</b> Sangat Baik. Realisasi melampaui target anggaran.<br><b>Rekomendasi:</b> Pertahankan kinerja pemungutan dan eksplorasi potensi pendapatan berkelanjutan."
        elif val >= 90:
            return "<b>Kondisi:</b> Baik. Realisasi hampir mencapai target.<br><b>Rekomendasi:</b> Tingkatkan pengawasan kepatuhan dan evaluasi pos pendapatan yang masih di bawah target."
        else:
            return "<b>Kondisi:</b> Kurang. Target pendapatan tidak tercapai secara signifikan.<br><b>Rekomendasi:</b> Lakukan intensifikasi dan ekstensifikasi penerimaan serta evaluasi penetapan target."

    elif "Efisiensi" in ratio_name:
        if val > 100:
            return "<b>Kondisi:</b> Kurang Efisien (Overbudget). Realisasi belanja melebihi pagu anggaran.<br><b>Rekomendasi:</b> Evaluasi penyebab kelebihan belanja, kendalikan pengeluaran non-prioritas, dan patuhi ketentuan anggaran."
        elif val >= 90:
            return "<b>Kondisi:</b> Sangat Efisien. Penyerapan anggaran sesuai target.<br><b>Rekomendasi:</b> Pastikan output dan manfaat program tetap optimal (Value for Money)."
        else:
            return "<b>Kondisi:</b> Penyerapan Rendah. Anggaran tidak terserap maksimal.<br><b>Rekomendasi:</b> Percepat jadwal pelaksanaan kegiatan/pengadaan dan lakukan monitoring berkala."

    elif "Kemandirian" in ratio_name:
        if val >= 50:
            return "<b>Kondisi:</b> Mandiri. PAD mendominasi struktur pendapatan.<br><b>Rekomendasi:</b> Jaga iklim investasi dan keberlanjutan sumber PAD."
        elif val >= 25:
            return "<b>Kondisi:</b> Sedang. Masih cukup bergantung pada dana transfer.<br><b>Rekomendasi:</b> Optimalkan potensi pajak/retribusi daerah dan layanan publik."
        else:
            return "<b>Kondisi:</b> Rendah (Sangat Tergantung Transfer).<br><b>Rekomendasi:</b> Perkuat basis penerimaan lokal dan efisiensi belanja."

    elif "Current Ratio" in ratio_name:
        if val >= 1.5:
            return "<b>Kondisi:</b> Likuiditas Sehat.<br><b>Rekomendasi:</b> Kelola kas surplus secara produktif dan pantau siklus modal kerja."
        else:
            return "<b>Kondisi:</b> Risiko Likuiditas.<br><b>Rekomendasi:</b> Percepat penagihan piutang, kelola utang jangka pendek, dan jaga arus kas operasi."

    elif "Debt to Equity" in ratio_name:
        if val > 2.0:
            return "<b>Kondisi:</b> Solvabilitas Berisiko (High Leverage).<br><b>Rekomendasi:</b> Kendalikan beban utang dan tingkatkan porsi pendanaan internal/ekuitas."
        else:
            return "<b>Kondisi:</b> Struktur Modal Sehat.<br><b>Rekomendasi:</b> Pertahankan keseimbangan utang dan ekuitas."

    return "<b>Analisis:</b> Detail belum tersedia untuk kondisi ini."

# ============================================================
# EXPORT FUNCTIONS (AUDIT & CASH FLOW)
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

def generate_cashflow_excel(cf_data, entity_type, filename):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_rows = [
            {"Indikator": "Entitas", "Nilai": entity_type},
            {"Indikator": "Nama File", "Nilai": filename},
            {"Indikator": "Total Kas Masuk (Cash Inflow)", "Nilai": cf_data["total_inflow"]},
            {"Indikator": "Total Kas Keluar (Cash Outflow)", "Nilai": cf_data["total_outflow"]},
            {"Indikator": "Arus Kas Bersih (Net Cash Flow)", "Nilai": cf_data["net_cf"]},
            {"Indikator": "Saldo Kas Awal", "Nilai": cf_data["opening_cash"]},
            {"Indikator": "Saldo Kas Akhir", "Nilai": cf_data["ending_cash"]},
            {"Indikator": "Perubahan vs Periode Sebelumnya", "Nilai": cf_data["cf_change_vs_prev"]},
            {"Indikator": "Arus Kas Operasi (PSAK 207)", "Nilai": cf_data["op_cf"]},
            {"Indikator": "Arus Kas Investasi (PSAK 207)", "Nilai": cf_data["inv_cf"]},
            {"Indikator": "Arus Kas Pendanaan (PSAK 207)", "Nilai": cf_data["fin_cf"]},
        ]
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Ringkasan Cash Flow", index=False)

        ratio_rows = [{"Rasio / Metrik": k, "Nilai": v} for k, v in cf_data["ratios"].items()]
        pd.DataFrame(ratio_rows).to_excel(writer, sheet_name="Rasio Arus Kas", index=False)

        if not cf_data["trend_df"].empty:
            cf_data["trend_df"].to_excel(writer, sheet_name="Tren Arus Kas", index=False)

        anom_rows = [{"Tingkat Risiko": r, "Anomali": t, "Detail": d, "Rekomendasi": rec} for r, t, d, rec in cf_data["anomalies"]]
        pd.DataFrame(anom_rows if anom_rows else [{"Status": "Tidak ada anomali terdeteksi"}]).to_excel(writer, sheet_name="Anomali & Risiko", index=False)
    buf.seek(0)
    return buf

def generate_cashflow_pdf(cf_data, entity_type, filename, ai_text=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph("LAPORAN ANALISIS CASH FLOW - FINORA AI", styles["Heading1"]))
    elements.append(Paragraph(f"Entitas: {entity_type} | File: {filename} | Tanggal: {datetime.now().strftime('%d %B %Y')}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    tbl_data = [
        ["Komponen Arus Kas", "Nilai"],
        ["Total Kas Masuk (Cash Inflow)", format_rupiah(cf_data["total_inflow"])],
        ["Total Kas Keluar (Cash Outflow)", format_rupiah(cf_data["total_outflow"])],
        ["Arus Kas Bersih (Net Cash Flow)", format_rupiah(cf_data["net_cf"])],
        ["Saldo Kas Awal", format_rupiah(cf_data["opening_cash"])],
        ["Saldo Kas Akhir", format_rupiah(cf_data["ending_cash"])],
        ["Arus Kas Operasi", format_rupiah(cf_data["op_cf"])],
        ["Arus Kas Investasi", format_rupiah(cf_data["inv_cf"])],
        ["Arus Kas Pendanaan", format_rupiah(cf_data["fin_cf"])],
    ]
    t = Table(tbl_data, colWidths=[260, 220])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7A1F3D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 14))

    elements.append(Paragraph("Jawaban Substantif Analisis Kas", styles["Heading2"]))
    elements.append(Paragraph(f"1. Dari mana kas diperoleh? {cf_data['sumber_kas']}", styles["Normal"]))
    elements.append(Paragraph(f"2. Ke mana kas digunakan? {cf_data['penggunaan_kas']}", styles["Normal"]))
    elements.append(Paragraph(f"3. Kondisi pengelolaan kas: {cf_data['kemampuan_kelola']}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    if ai_text:
        elements.append(Paragraph("AI Financial Insights & Rekomendasi", styles["Heading2"]))
        clean_ai = re.sub(r"<[^>]+>", " ", ai_text)
        elements.append(Paragraph(clean_ai[:2200], styles["Normal"]))

    doc.build(elements)
    buf.seek(0)
    return buf

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
    uploaded = c1.file_uploader("Upload File Excel/CSV/PDF", type=["xlsx", "csv", "pdf"], key="audit_upload_unique")
    header_row = c2.number_input("Header di baris ke-", min_value=0, value=0, key="audit_header_unique")

    if uploaded:
        df = read_uploaded_dataframe(uploaded, header_row=header_row)
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
    finora_header("Analisis Laporan Keuangan", "Evaluasi Kesehatan Finansial Otomatis")
    st.markdown("### 1. Upload Laporan Keuangan")
    c1, c2, c3 = st.columns([2, 1, 1])
    uploaded_fin = c1.file_uploader("Upload File Excel/CSV/PDF (APBD / Neraca / Laba Rugi)", type=["xlsx", "csv", "pdf"], key="fin_upload_unique")
    entity_type = c2.selectbox("Jenis Entitas", ["Perusahaan Swasta / Tbk", "Pemerintahan / BLU"], key="fin_entity_unique")
    target_year = c3.selectbox("Periode / Tahun Acuan", ["2024", "2023"], key="fin_year_unique")

    if uploaded_fin:
        df_fin = read_uploaded_dataframe(uploaded_fin, header_row=0)
        st.success("✨ File berhasil dibaca. Klik tombol di bawah untuk analisis keuangan otomatis.")
        with st.expander("📄 Pratinjau Data Terbaca", expanded=False):
            st.dataframe(df_fin.head(25), use_container_width=True)

        if st.button("📊 ANALISIS KEUANGAN OTOMATIS", use_container_width=True):
            with st.spinner("Menghitung rasio..."):
                nilai_utama = auto_extract_financials(df_fin, entity_type, target_year=target_year)
                rasio = hitung_rasio_dari_nilai(nilai_utama, entity_type)
                st.session_state['fin_results'] = {
                    "rasio": rasio, "nilai": nilai_utama, "entity_type": entity_type, "filename": uploaded_fin.name, "target_year": target_year
                }

    if 'fin_results' in st.session_state:
        res = st.session_state['fin_results']
        rasio, nilai = res["rasio"], res["nilai"]
        st.markdown("---")
        st.markdown(f"### Hasil Analisis: **{res['entity_type']}** (Tahun {res.get('target_year', '2024')})")
        cols = st.columns(len(rasio) if len(rasio) > 0 else 1)
        for i, (k, v) in enumerate(rasio.items()):
            with cols[i % len(cols)]:
                metric_card(k, f"{v}")

        st.markdown("### 💡 Insight & Rekomendasi Rasio")
        for k, v in rasio.items():
            st.markdown(f'<div class="insight-box"><h4>{k}: {v}</h4><p>{get_detailed_insight(k, v)}</p></div>', unsafe_allow_html=True)

def page_cash_flow():
    finora_header("FINORA AI – Cash Flow Analysis", "Analisis substantif arus kas: sumber kas, penggunaan kas, rasio, tren, anomali & AI insight")
    st.markdown("### 1. Upload Data Transaksi atau Laporan Arus Kas")
    c1, c2, c3 = st.columns([2, 1, 1])
    uploaded_cf = c1.file_uploader("Upload File Excel/CSV/PDF (Mutasi Kas / LRA / LAK / Laporan Arus Kas)", type=["xlsx", "csv", "pdf"], key="cf_menu_upload")
    entity_type = c2.selectbox("Jenis Entitas", ["Perusahaan Swasta / Tbk", "Pemerintahan / BLU"], key="cf_menu_entity")
    target_year = c3.selectbox("Tahun Analisis", ["2024", "2023"], key="cf_menu_year")

    c_opt1, c_opt2 = st.columns(2)
    opening_cash_input = c_opt1.number_input("Saldo Kas Awal (Opsional / Rp — 0 untuk baca otomatis dari file)", value=0.0, step=1000000.0, key="cf_opening_cash")
    prev_net_cf_input = c_opt2.number_input("Net Cash Flow Periode Sebelumnya (Opsional / Rp — 0 untuk baca otomatis)", value=0.0, step=1000000.0, key="cf_prev_net_cf")

    if uploaded_cf:
        df_cf = read_uploaded_dataframe(uploaded_cf, header_row=0)
        preview_txt = extract_text_preview(uploaded_cf, df_cf)
        st.success("✨ File berhasil dibaca. Klik tombol di bawah untuk menjalankan analisis cash flow substantif.")
        with st.expander("📄 Pratinjau Baris Terbaca (Termasuk 2 Kolom Tahun jika PDF)", expanded=False):
            st.dataframe(df_cf.head(30), use_container_width=True)

        if st.button("💵 ANALISIS CASH FLOW OTOMATIS", use_container_width=True):
            with st.spinner("Menganalisis arus kas, rasio, tren, dan anomali..."):
                cf_data = analyze_cash_flow_substantive(
                    df_cf,
                    entity_type=entity_type,
                    opening_cash=opening_cash_input,
                    prev_net_cf=prev_net_cf_input,
                    target_year=target_year,
                )
                st.session_state["cf_menu_results"] = {
                    "cf": cf_data,
                    "entity_type": entity_type,
                    "filename": uploaded_cf.name,
                    "preview_txt": preview_txt,
                    "target_year": target_year,
                }

    if "cf_menu_results" in st.session_state:
        res_cf = st.session_state["cf_menu_results"]
        cf = res_cf["cf"]
        st.markdown("---")

        st.markdown(f"### 1. Cash Flow Overview (Ringkasan Arus Kas — Tahun {res_cf.get('target_year', '2024')})")
        ov1, ov2, ov3, ov4 = st.columns(4)
        with ov1:
            metric_card("Total Cash Inflow", format_rupiah(cf["total_inflow"]))
        with ov2:
            metric_card("Total Cash Outflow", format_rupiah(cf["total_outflow"]))
        with ov3:
            metric_card("Net Cash Flow", format_rupiah(cf["net_cf"]))
        with ov4:
            metric_card("Ending Cash Balance", format_rupiah(cf["ending_cash"]))

        ov5, ov6 = st.columns(2)
        with ov5:
            metric_card("Saldo Kas Awal", format_rupiah(cf["opening_cash"]))
        with ov6:
            pct_str = f" ({cf['cf_change_pct']:+.2f}%)" if cf["cf_change_pct"] != 0 else ""
            metric_card("Perubahan vs Periode Lalu", f"{format_rupiah(cf['cf_change_vs_prev'])}{pct_str}")

        st.markdown(f"""
        <div class="insight-box">
            <h4>Jawaban Substantif Pengelolaan Kas</h4>
            <p><b>Dari mana kas diperoleh?</b> {cf['sumber_kas']}</p>
            <p><b>Ke mana kas digunakan?</b> {cf['penggunaan_kas']}</p>
            <p><b>Bagaimana kondisi kemampuan {cf.get('entity_label', 'perusahaan')} dalam mengelola kasnya?</b> {cf['kemampuan_kelola']}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 2. Operating, Investing & Financing Analysis (PSAK 207 / SAP)")
        act1, act2, act3 = st.columns(3)
        with act1:
            metric_card("Arus Kas Operasi", format_rupiah(cf["op_cf"]))
        with act2:
            metric_card("Arus Kas Investasi", format_rupiah(cf["inv_cf"]))
        with act3:
            metric_card("Arus Kas Pendanaan", format_rupiah(cf["fin_cf"]))

        act_rows = pd.DataFrame([
            {"Aktivitas": "Arus Kas Operasi (Operating)", "Nilai": format_rupiah(cf["op_cf"]), "Fokus Evaluasi": "Penerimaan operasional/PAD/transfer, belanja pegawai & barang/jasa, pembayaran pemasok, serta biaya operasional inti."},
            {"Aktivitas": "Arus Kas Investasi (Investing)", "Nilai": format_rupiah(cf["inv_cf"]), "Fokus Evaluasi": "Belanja modal, pembelian/penjualan aset tetap, dan instrumen investasi."},
            {"Aktivitas": "Arus Kas Pendanaan / Pembiayaan", "Nilai": format_rupiah(cf["fin_cf"]), "Fokus Evaluasi": "Pembiayaan netto, pinjaman, pembayaran utang, setoran modal, dan dividen."},
        ])
        st.table(act_rows)

        st.markdown("### 3. Cash Flow Ratios (Rasio Arus Kas)")
        r_cols = st.columns(3)
        ratio_items = list(cf["ratios"].items())
        for idx, (r_name, r_val) in enumerate(ratio_items):
            with r_cols[idx % 3]:
                formatted_r = format_rupiah(r_val) if "FCF" in r_name else f"{r_val}"
                metric_card(r_name, formatted_r)

        st.markdown("""
        <div class="insight-box">
            <h4>Panduan Interpretasi Rasio Arus Kas</h4>
            <p><b>Operating Cash Flow Ratio:</b> Mengukur kemampuan kas operasi membayar kewajiban jangka pendek.</p>
            <p><b>Cash Flow to Debt Ratio:</b> Mengukur daya dukung kas operasi terhadap total liabilitas/utang.</p>
            <p><b>Cash Flow Margin:</b> Mengukur efisiensi konversi pendapatan/penerimaan menjadi kas operasi.</p>
            <p><b>Free Cash Flow (FCF):</b> Sisa kas setelah pengeluaran investasi/belanja modal untuk ekspansi atau pelunasan kewajiban.</p>
            <p><b>Cash Flow Coverage:</b> Mengukur kemampuan kas operasi menutup kewajiban pokok/pendanaan.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 4. Cash Flow Trends (Tren Arus Kas)")
        trend_df = cf["trend_df"]
        fig_trend = None
        if not trend_df.empty:
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(x=trend_df["Periode"], y=trend_df["Kas Masuk"], name="Kas Masuk"))
            fig_trend.add_trace(go.Bar(x=trend_df["Periode"], y=trend_df["Kas Keluar"], name="Kas Keluar"))
            fig_trend.add_trace(go.Scatter(x=trend_df["Periode"], y=trend_df["Saldo Kas Kumulatif"], name="Saldo Kas Kumulatif", mode="lines+markers"))
            fig_trend.update_layout(barmode="group", title="Tren Kas Masuk, Kas Keluar, dan Saldo Kas", height=380)
            st.plotly_chart(fig_trend, use_container_width=True)
            st.dataframe(trend_df, use_container_width=True)

        st.markdown("### 5. Anomaly & Risk Detection (Deteksi Masalah & Anomali)")
        if cf["anomalies"]:
            for risk_lvl, title, detail, rec in cf["anomalies"]:
                with st.expander(f"[{risk_lvl}] {title}"):
                    st.markdown(f"**Temuan:** {detail}")
                    st.error(f"**Rekomendasi:** {rec}")
        else:
            st.success("Tidak terdeteksi anomali atau risiko kas yang signifikan pada data yang dianalisis.")

        st.markdown("### 6. AI Financial Insights")
        default_ai_summary = (
            f"<b>Ringkasan:</b> Arus kas bersih tercatat {format_rupiah(cf['net_cf'])} dengan arus kas operasi {format_rupiah(cf['op_cf'])} "
            f"dan saldo kas akhir {format_rupiah(cf['ending_cash'])}.<br>"
            f"<b>Temuan Utama:</b> {cf['sumber_kas']} {cf['penggunaan_kas']}<br>"
            f"<b>Rekomendasi:</b> {cf['kemampuan_kelola']} Lakukan pemantauan arus kas masuk/keluar secara berkala dan tindak lanjuti temuan anomali."
        )
        st.markdown(f'<div class="insight-box"><h4>FINORA AI – Cash Flow Insights</h4><p>{default_ai_summary}</p></div>', unsafe_allow_html=True)

        if st.button("✨ Generate Analisis Naratif AI Mendalam", use_container_width=True, key="cf_ai_btn"):
            ctx_cf = (
                f"Entitas: {res_cf['entity_type']} | File: {res_cf['filename']} | Tahun: {res_cf.get('target_year', '2024')}\n"
                f"Cash Inflow: {cf['total_inflow']}, Cash Outflow: {cf['total_outflow']}, Net CF: {cf['net_cf']}, "
                f"Saldo Awal: {cf['opening_cash']}, Saldo Akhir: {cf['ending_cash']}\n"
                f"Operasi: {cf['op_cf']}, Investasi: {cf['inv_cf']}, Pendanaan: {cf['fin_cf']}\n"
                f"Rasio: {cf['ratios']}\n"
                f"Anomali: {[a[1] for a in cf['anomalies']]}\n"
                f"Cuplikan Data: {res_cf.get('preview_txt', '')[:1000]}"
            )
            prompt_cf = (
                "Susun analisis Cash Flow substantif dalam Bahasa Indonesia dengan bagian: "
                "1) Ringkasan & Sumber/Penggunaan Kas, 2) Evaluasi Aktivitas & Rasio, "
                "3) Risiko/Anomali Utama, 4) Rekomendasi Tindak Lanjut."
            )
            ai_out = st.write_stream(ask_gemini_stream(prompt_cf, ctx_cf, system_role="Analis Keuangan & Auditor Cash Flow"))
            st.session_state["cf_ai_report_text"] = ai_out

        st.markdown("### 7. Export Report (Laporan & Hasil Analisis)")
        ai_saved = st.session_state.get("cf_ai_report_text", default_ai_summary)
        exp1, exp2, exp3 = st.columns(3)
        with exp1:
            cf_xlsx = generate_cashflow_excel(cf, res_cf["entity_type"], res_cf["filename"])
            st.download_button("⬇️ Download Excel (.xlsx)", data=cf_xlsx, file_name=f"CashFlow_{res_cf['filename']}.xlsx", use_container_width=True)
        with exp2:
            cf_pdf = generate_cashflow_pdf(cf, res_cf["entity_type"], res_cf["filename"], ai_text=ai_saved)
            st.download_button("⬇️ Download PDF (.pdf)", data=cf_pdf, file_name=f"CashFlow_{res_cf['filename']}.pdf", use_container_width=True)
        with exp3:
            ai_report_txt = (
                f"LAPORAN AI CASH FLOW - {res_cf['entity_type']} ({res_cf['filename']})\n"
                f"Tanggal: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
                f"Sumber Kas: {cf['sumber_kas']}\n"
                f"Penggunaan Kas: {cf['penggunaan_kas']}\n"
                f"Kemampuan Pengelolaan Kas: {cf['kemampuan_kelola']}\n\n"
                f"{re.sub(r'<[^>]+>', ' ', ai_saved)}"
            )
            st.download_button("⬇️ Download AI Report (.txt)", data=ai_report_txt.encode("utf-8"), file_name=f"AI_CashFlow_Report_{res_cf['filename']}.txt", use_container_width=True)

        if fig_trend is not None:
            st.markdown("#### Unduh Grafik Tren Arus Kas")
            try:
                png_bytes = fig_trend.to_image(format="png", scale=2)
                st.download_button("⬇️ Download Grafik Tren (.png)", data=png_bytes, file_name=f"Grafik_CashFlow_{res_cf['filename']}.png", use_container_width=True)
            except Exception:
                st.caption("Tip: Install `kaleido` (`pip install kaleido`) untuk mengunduh grafik langsung dalam format PNG.")

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
        if an:
            summary_data = json.loads(an["summary_json"])
            entitas = summary_data.get("entitas", "-")
            periode = summary_data.get("periode", "-")
            score = summary_data.get("score", 0)
            flagged_data = {k: pd.DataFrame(v) for k, v in summary_data.get("flagged_data", {}).items()}
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
            Platform Kecerdasan Finansial untuk mendeteksi anomali, menyusun Kertas Kerja Audit (KKA), serta Analisis Laporan Keuangan secara otomatis.
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
            ["Dashboard", "Audit Profesional", "Analisis Laporan Keuangan", "Analisis Cash Flow", "Pusat Laporan & Arsip", "Kembali ke Beranda"]
        )

        if menu == "Dashboard":
            page_dashboard()
        elif menu == "Audit Profesional":
            page_audit()
        elif menu == "Analisis Laporan Keuangan":
            page_financial_analysis()
        elif menu == "Analisis Cash Flow":
            page_cash_flow()
        elif menu == "Pusat Laporan & Arsip":
            page_simpan()
        elif menu == "Kembali ke Beranda":
            st.session_state.app_started = False
            st.rerun()

if __name__ == "__main__":
    main()
