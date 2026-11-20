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
