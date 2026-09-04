import streamlit as st
import sqlite3
import bcrypt
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import urllib.parse
import os

# --- KONFIGURASI AI (GOOGLE GEMINI) ---
# Masukkan API Key Gemini Anda di sini. Dapatkan di: https://aistudio.google.com/
GEMINI_API_KEY = "MASUKKAN_API_KEY_GEMINI_ANDA_DISINI" 

try:
    import google.generativeai as genai
    genai.configure(api_key=GEMINI_API_KEY)
    AI_AVAILABLE = True
except ImportError:
    AI_AVAILABLE = False

# ==========================================
# KONFIGURASI HALAMAN & CSS
# ==========================================
st.set_page_config(page_title="Monitoring Binpres KONI", page_icon="🏆", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .main-header { font-size: 38px; font-weight: 800; color: #1E3A8A; text-align: center; margin-bottom: -10px; }
    .sub-header { font-size: 18px; font-weight: 400; color: #64748B; text-align: center; margin-bottom: 30px; }
    .divider { height: 2px; background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%); margin: 20px 0; }
    .news-card { background-color: #F8FAFC; padding: 20px; border-radius: 10px; border-left: 5px solid #10B981; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
    .news-date { color: #64748B; font-size: 12px; margin-bottom: 5px; }
    .news-title { font-size: 20px; font-weight: bold; color: #0F172A; margin-bottom: 10px; }
    .schedule-card { background-color: #ffffff; border: 1px solid #E2E8F0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); height: 100%;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# DATABASE
# ==========================================
DB_NAME = "monitoring.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT, role TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS schedules (id INTEGER PRIMARY KEY AUTOINCREMENT, cabor TEXT, tanggal TEXT, tempat TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS reports (id INTEGER PRIMARY KEY AUTOINCREMENT, cabor TEXT, tanggal_kegiatan TEXT, submit_time DATETIME, file_name TEXT, file_data BLOB)''')
    # Tabel baru untuk berita AI
    c.execute('''CREATE TABLE IF NOT EXISTS news (id INTEGER PRIMARY KEY AUTOINCREMENT, cabor TEXT, tanggal TEXT, judul TEXT, konten TEXT, created_at DATETIME)''')

    c.execute("DELETE FROM reports WHERE submit_time < datetime('now', '-30 days')")

    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        hashed_pw = bcrypt.hashpw('admin123'.encode('utf-8'), bcrypt.gensalt())
        c.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", ('admin', hashed_pw, 'admin'))
    conn.commit()
    conn.close()

def authenticate(username, password):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT password, role FROM users WHERE username=?", (username,))
    result = c.fetchone()
    conn.close()
    if result and bcrypt.checkpw(password.encode('utf-8'), result[0]):
        return True, result[1]
    return False, None

def get_current_time_id():
    now = datetime.datetime.now()
    hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"]
    return f"{hari[now.weekday()]}, {now.day} {bulan[now.month - 1]} {now.year}"

# ==========================================
# FUNGSI GENERATE BERITA AI
# ==========================================
def generate_news_with_ai(cabor, tanggal, tempat, catatan):
    if not AI_AVAILABLE or GEMINI_API_KEY == "MASUKKAN_API_KEY_GEMINI_ANDA_DISINI":
        # Fallback jika API Key belum disetting
        judul = f"Update Monev: Evaluasi Cabor {cabor} Berjalan Lancar"
        konten = f"Kegiatan monitoring cabang olahraga {cabor} telah dilaksanakan pada {tanggal} di {tempat.title()}. Laporan terbaru menunjukkan evaluasi mendalam untuk peningkatan performa atlet kedepannya. Evaluasi ini diharapkan mampu mendongkrak prestasi menuju target yang dicanangkan."
        return judul, konten

    try:
        model = genai.GenerativeModel('gemini-pro')
        prompt = f"Buatkan 1 paragraf berita olahraga yang formal, jurnalistik, dan memotivasi (maks 80 kata) berdasarkan hasil monitoring berikut.\nCabor: {cabor}\nTanggal: {tanggal}\nTempat: {tempat}\nCatatan Evaluasi: {catatan}\n\nFormat keluaran WAJIB 2 baris:\nBaris 1: Judul Berita (catchy)\nBaris 2: Isi berita."
        response = model.generate_content(prompt)
        teks = response.text.split('\n', 1)
        
        judul = teks[0].replace("**", "").replace("Judul:", "").strip()
        konten = teks[1].replace("**", "").strip() if len(teks) > 1 else "Berita sedang disiapkan."
        return judul, konten
    except Exception as e:
        return f"Update Monev {cabor}", f"Monitoring telah selesai dilaksanakan di {tempat}. AI gagal memproses detail berita: {e}"

# ==========================================
# FUNGSI GENERATE WORD
# ==========================================
def generate_word_report(petugas_text, cabor, tanggal, tempat, catatan, fotos):
    doc = Document()
    head = doc.add_heading('LAPORAN MONITORING CABANG OLAHRAGA', 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Cabang Olahraga\t: {cabor}\nTanggal\t\t: {tanggal}\nTempat\t\t: {tempat}\n")
    
    doc.add_heading('Daftar Petugas:', level=3)
    for i, p in enumerate([p.strip() for p in petugas_text.split('\n') if p.strip()], 1):
        doc.add_paragraph(f"{i}. {p}")

    doc.add_heading('\nCatatan Evaluasi / Hasil Monitoring:', level=3)
    doc.add_paragraph(catatan)

    if fotos:
        doc.add_page_break()
        doc.add_heading('Lampiran Foto Dokumentasi', level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        table = doc.add_table(rows=0, cols=2)
        row_cells = None
        for idx, foto in enumerate(fotos):
            if idx % 2 == 0: row_cells = table.add_row().cells
            p = row_cells[idx % 2].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                p.add_run().add_picture(io.BytesIO(foto.getvalue()), width=Inches(2.8)) 
            except:
                p.add_run().add_text("(Gagal memuat gambar)")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# STATE ROUTING UTAMA
# ==========================================
init_db()

if 'page' not in st.session_state: st.session_state['page'] = 'home'
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'selected_schedule' not in st.session_state: st.session_state['selected_schedule'] = None

# Navigasi Sidebar
st.sidebar.markdown("### 🏆 PANEL NAVIGASI")
st.sidebar.caption(get_current_time_id())

if st.sidebar.button("🏠 Halaman Depan", use_container_width=True):
    st.session_state['page'] = 'home'
    st.session_state['selected_schedule'] = None
    st.rerun()

if st.session_state['logged_in']:
    st.sidebar.success(f"👤 Login: {st.session_state['username']}")
    if st.session_state['role'] == 'admin':
        if st.sidebar.button("📅 Kelola Jadwal", use_container_width=True): st.session_state['page'] = 'admin_jadwal'; st.rerun()
        if st.sidebar.button("📂 Arsip Laporan", use_container_width=True): st.session_state['page'] = 'admin_arsip'; st.rerun()
        if st.sidebar.button("👥 Kelola User", use_container_width=True): st.session_state['page'] = 'admin_user'; st.rerun()
    if st.sidebar.button("🚪 Keluar (Logout)", use_container_width=True, type="secondary"):
        st.session_state['logged_in'] = False
        st.session_state['page'] = 'home'
        st.rerun()
else:
    if st.sidebar.button("🔐 Login Admin/Petugas", use_container_width=True):
        st.session_state['page'] = 'login'
        st.rerun()

st.sidebar.markdown("---")

# ==========================================
# VIEW: HALAMAN DEPAN (PUBLIK)
# ==========================================
if st.session_state['page'] == 'home':
    st.markdown("<div class='main-header'>🏆 E-MONEV CABOR</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Binpres KONI Kabupaten Tangerang</div>", unsafe_allow_html=True)
    
    # BAGIAN 1: JADWAL MONITORING
    st.markdown("### 📅 Jadwal Monitoring Mendatang")
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, cabor, tanggal, tempat FROM schedules ORDER BY id DESC")
    schedules = c.fetchall()
    
    if not schedules:
        st.info("Belum ada jadwal monitoring yang aktif saat ini.")
    else:
        cols = st.columns(3)
        for idx, s in enumerate(schedules):
            with cols[idx % 3]:
                st.markdown(f"""
                <div class='schedule-card'>
                    <h4 style='color:#1E3A8A; margin-bottom:5px;'>{s[1]}</h4>
                    <p style='margin:0; font-size:14px; color:#64748B;'>📍 {s[3]}</p>
                    <p style='margin:0; font-size:14px; font-weight:bold;'>🗓️ {s[2]}</p>
                </div>
                <br>
                """, unsafe_allow_html=True)
                if st.button(f"📝 Isi Laporan", key=f"btn_jadwal_{s[0]}", use_container_width=True):
                    st.session_state['selected_schedule'] = s
                    st.session_state['page'] = 'form'
                    st.rerun()
    
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    
    # BAGIAN 2: NEWS FEED AI
    st.markdown("### 📰 Berita Terkini (AI Generated)")
    c.execute("SELECT cabor, tanggal, judul, konten, created_at FROM news ORDER BY created_at DESC LIMIT 5")
    news_data = c.fetchall()
    conn.close()
    
    if not news_data:
        st.caption("Belum ada laporan yang disubmit. Berita akan muncul otomatis setelah laporan dibuat.")
    else:
        for n in news_data:
            st.markdown(f"""
            <div class='news-card'>
                <div class='news-date'>{n[1]} • Waktu Rilis: {n[4]}</div>
                <div class='news-title'>{n[2]}</div>
                <div style='color: #334155; line-height: 1.6;'>{n[3]}</div>
            </div>
            """, unsafe_allow_html=True)

# ==========================================
# VIEW: LOGIN
# ==========================================
elif st.session_state['page'] == 'login':
    col1, col2, col3 = st.columns([1.5, 2, 1.5])
    with col2:
        with st.container(border=True):
            st.markdown("#### 🔐 Silakan Masuk")
            with st.form("login_form"):
                user_in = st.text_input("👤 Username")
                pass_in = st.text_input("🔑 Password", type="password")
                if st.form_submit_button("Masuk", use_container_width=True):
                    is_auth, role = authenticate(user_in, pass_in)
                    if is_auth:
                        st.session_state['logged_in'] = True
                        st.session_state['username'] = user_in
                        st.session_state['role'] = role
                        st.session_state['page'] = 'form' if st.session_state.get('selected_schedule') else 'home'
                        st.rerun()
                    else:
                        st.error("🚨 Username/Password salah!")

# ==========================================
# VIEW: FORM LAPORAN
# ==========================================
elif st.session_state['page'] == 'form':
    if not st.session_state['logged_in']:
        st.warning("⚠️ Anda harus login terlebih dahulu untuk mengisi laporan jadwal ini.")
        st.session_state['page'] = 'login'
        st.rerun()
    
    if not st.session_state.get('selected_schedule'):
        st.error("Jadwal belum dipilih. Kembali ke Halaman Depan.")
        st.button("Kembali", on_click=lambda: st.session_state.update({'page': 'home'}))
    else:
        s_id, val_cabor, val_tanggal, val_tempat = st.session_state['selected_schedule']
        
        st.markdown(f"### 📝 Form Laporan Monitoring")
        st.info(f"**Target Cabor:** {val_cabor} | **Lokasi:** {val_tempat} | **Tanggal:** {val_tanggal}")
        
        with st.container(border=True):
            petugas_text = st.text_area("👤 Daftar Petugas (1 nama per baris)", height=100)
            catatan = st.text_area("✍️ Catatan Evaluasi (Akan dirangkum AI menjadi Berita)", height=150)
            fotos = st.file_uploader("📸 Upload Bukti (2-5 Foto)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

            if st.button("📄 Submit Laporan & Generate Berita", use_container_width=True, type="primary"):
                if not petugas_text.strip() or not catatan.strip() or len(fotos) < 2 or len(fotos) > 5:
                    st.error("⚠️ Lengkapi data! Catatan wajib diisi, dan unggah 2-5 foto.")
                else:
                    with st.spinner("⏳ Memproses Dokumen & AI merekap berita..."):
                        # 1. Generate Word Document
                        word_file = generate_word_report(petugas_text, val_cabor, val_tanggal, val_tempat, catatan, fotos)
                        safe_date = val_tanggal.replace(" s/d ", "_").replace("-", "").replace("/", "")
                        file_name_doc = f"Monev_{val_cabor.split()[0]}_{safe_date}.docx"
                        
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        conn = sqlite3.connect(DB_NAME)
                        c = conn.cursor()
                        
                        # Simpan Dokumen
                        c.execute("INSERT INTO reports (cabor, tanggal_kegiatan, submit_time, file_name, file_data) VALUES (?, ?, ?, ?, ?)",
                                  (val_cabor, val_tanggal, now_str, file_name_doc, word_file.getvalue()))
                        
                        # 2. Generate Berita AI & Simpan
                        judul_ai, konten_ai = generate_news_with_ai(val_cabor, val_tanggal, val_tempat, catatan)
                        c.execute("INSERT INTO news (cabor, tanggal, judul, konten, created_at) VALUES (?, ?, ?, ?, ?)",
                                  (val_cabor, val_tanggal, judul_ai, konten_ai, now_str))
                        
                        conn.commit()
                        conn.close()
                        
                        st.session_state['report_done'] = True
                        st.session_state['word_file'] = word_file
                        st.session_state['doc_name'] = file_name_doc
        
        if st.session_state.get('report_done'):
            st.success("🎉 Laporan berhasil disimpan & Berita Otomatis telah terbit di Halaman Depan!")
            colA, colB = st.columns(2)
            with colA:
                st.download_button("📥 Unduh Salinan Word", data=st.session_state['word_file'], file_name=st.session_state['doc_name'], use_container_width=True)
            with colB:
                if st.button("🏠 Kembali ke Beranda", use_container_width=True):
                    st.session_state['report_done'] = False
                    st.session_state['page'] = 'home'
                    st.rerun()

# ==========================================
# VIEW: ADMIN (ARSIP, JADWAL, USER)
# ==========================================
elif st.session_state['page'].startswith('admin_') and st.session_state['role'] == 'admin':
    
    if st.session_state['page'] == 'admin_jadwal':
        st.markdown("### 📅 Kelola Jadwal Monitoring")
        conn = sqlite3.connect(DB_NAME)
        
        with st.form("form_jadwal"):
            cabor_in = st.text_input("Nama Cabor (Contoh: PANAHAN)")
            tgl_in = st.text_input("Tanggal Kegiatan (Contoh: 12-10-2024)")
            tmpt_in = st.text_input("Lokasi")
            if st.form_submit_button("Tambah Jadwal"):
                conn.cursor().execute("INSERT INTO schedules (cabor, tanggal, tempat) VALUES (?, ?, ?)", (cabor_in.upper(), tgl_in, tmpt_in))
                conn.commit()
                st.rerun()
                
        df_jadwal = pd.read_sql_query("SELECT id, cabor, tanggal, tempat FROM schedules", conn)
        st.dataframe(df_jadwal, use_container_width=True, hide_index=True)
        del_id = st.selectbox("Hapus ID:", df_jadwal['id'].tolist() if not df_jadwal.empty else ["Kosong"])
        if st.button("Hapus Jadwal") and del_id != "Kosong":
            conn.cursor().execute("DELETE FROM schedules WHERE id=?", (del_id,))
            conn.commit()
            st.rerun()
        conn.close()

    elif st.session_state['page'] == 'admin_arsip':
        st.markdown("### 📂 Arsip Laporan & Berita AI")
        conn = sqlite3.connect(DB_NAME)
        df_reports = pd.read_sql_query("SELECT id, cabor, tanggal_kegiatan, submit_time, file_name FROM reports", conn)
        st.dataframe(df_reports, use_container_width=True, hide_index=True)
        
        dl_id = st.selectbox("Unduh/Hapus Laporan ID:", df_reports['id'].tolist() if not df_reports.empty else ["Kosong"])
        col1, col2 = st.columns(2)
        if dl_id != "Kosong":
            row = conn.cursor().execute("SELECT file_name, file_data FROM reports WHERE id=?", (dl_id,)).fetchone()
            if row:
                col1.download_button(label="📥 Unduh Dokumen", data=row[1], file_name=row[0], use_container_width=True)
            if col2.button("🗑️ Hapus Data", use_container_width=True):
                conn.cursor().execute("DELETE FROM reports WHERE id=?", (dl_id,))
                conn.commit()
                st.rerun()
        conn.close()
        
    elif st.session_state['page'] == 'admin_user':
        st.markdown("### 👥 Manajemen Pengguna")
        conn = sqlite3.connect(DB_NAME)
        st.dataframe(pd.read_sql_query("SELECT username, role FROM users", conn), use_container_width=True)
        
        with st.form("add_user"):
            new_u = st.text_input("Username Baru")
            new_p = st.text_input("Password", type="password")
            new_r = st.selectbox("Role", ["user", "admin"])
            if st.form_submit_button("Tambah"):
                hashed = bcrypt.hashpw(new_p.encode('utf-8'), bcrypt.gensalt())
                try:
                    conn.cursor().execute("INSERT INTO users VALUES (?, ?, ?)", (new_u, hashed, new_r))
                    conn.commit()
                    st.rerun()
                except: st.error("Username sudah ada!")
        conn.close()
