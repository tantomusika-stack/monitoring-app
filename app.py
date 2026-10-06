import streamlit as st
import sqlite3
import hashlib
from docx import Document
from docx.shared import Inches
import io
from datetime import datetime
from PIL import Image

# ==========================
# 1. SETUP DATABASE (SQLite)
# ==========================
conn = sqlite3.connect('users.db', check_same_thread=False)
c = conn.cursor()

def init_db():
    c.execute('CREATE TABLE IF NOT EXISTS users (username TEXT, password TEXT, role TEXT)')
    conn.commit()
    # Buat default admin jika belum ada
    c.execute('SELECT * FROM users WHERE username="admin"')
    if not c.fetchone():
        add_userdata("admin", make_hashes("admin123"), "admin")

def make_hashes(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_hashes(password, hashed_text):
    if make_hashes(password) == hashed_text:
        return True
    return False

def add_userdata(username, password, role):
    c.execute('INSERT INTO users(username, password, role) VALUES (?,?,?)', (username, password, role))
    conn.commit()

def login_user(username, password):
    c.execute('SELECT * FROM users WHERE username =? AND password = ?', (username, password))
    data = c.fetchall()
    return data

init_db()

# ==========================
# 2. FUNGSI GENERATE WORD
# ==========================
def generate_word_report(data, image_file, realtime_stamp):
    doc = Document()
    doc.add_heading('LAPORAN MONITORING CABANG OLAHRAGA', 0)
    
    # Informasi Dasar
    doc.add_heading('1. Informasi Kegiatan', level=1)
    doc.add_paragraph(f"Nama Petugas   : {data['staff']}")
    doc.add_paragraph(f"Pengurus       : {data['pengurus']}")
    doc.add_paragraph(f"Cabang Olahraga: {data['cabor']}")
    doc.add_paragraph(f"Tanggal        : {data['tanggal']}")
    doc.add_paragraph(f"Tempat         : {data['tempat']}")
    doc.add_paragraph(f"Nama Kegiatan  : {data['kegiatan']}")
    
    # Isi Laporan
    doc.add_heading('2. Hasil Evaluasi & Monitoring', level=1)
    for section, points in data['laporan'].items():
        doc.add_heading(section, level=2)
        for key, val in points.items():
            p = doc.add_paragraph()
            p.add_run(f"{key}: ").bold = True
            p.add_run(val)
            
    # Dokumentasi
    doc.add_heading('3. Dokumentasi Kegiatan', level=1)
    if image_file is not None:
        # Resize/Read image
        img = Image.open(image_file)
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        
        doc.add_picture(img_byte_arr, width=Inches(5))
        doc.add_paragraph(f"Waktu Pengambilan/Upload Foto: {realtime_stamp}")
    else:
        doc.add_paragraph("Tidak ada foto yang dilampirkan.")

    # Simpan ke BytesIO agar bisa didownload
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==========================
# 3. ANTARMUKA STREAMLIT
# ==========================
def main():
    st.set_page_config(page_title="App Monitoring Cabor", layout="wide")
    
    # Session State untuk Login
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""

    if not st.session_state["logged_in"]:
        st.title("Login Sistem Monitoring")
        username = st.text_input("Username")
        password = st.text_input("Password", type='password')
        if st.button("Login"):
            hashed_pswd = make_hashes(password)
            result = login_user(username, hashed_pswd)
            if result:
                st.session_state["logged_in"] = True
                st.session_state["username"] = username
                st.session_state["role"] = result[0][2]
                st.rerun()
            else:
                st.error("Username atau Password salah")
        return

    # Jika sudah login
    st.sidebar.title(f"Halo, {st.session_state['username']}")
    st.sidebar.write(f"Role: {st.session_state['role'].upper()}")
    
    menu = ["Form Monitoring"]
    if st.session_state["role"] == "admin":
        menu.append("Tambah User")
    menu.append("Logout")
    
    choice = st.sidebar.selectbox("Navigasi", menu)
    
    if choice == "Form Monitoring":
        st.title("Form Laporan Monitoring Cabang Olahraga")
        
        with st.form("form_monitoring"):
            col1, col2 = st.columns(2)
            with col1:
                staff = st.text_input("1. Nama Petugas Staff")
                pengurus_input = st.text_area("2. Daftar Pengurus (Pisahkan dengan koma, Maks 5 orang)")
                cabor = st.text_input("3. Cabang Olahraga")
                tanggal = st.date_input("4. Tanggal Monitoring")
            with col2:
                tempat = st.text_input("5. Tempat Monitoring")
                kegiatan = st.selectbox("6. Nama Kegiatan", ["TRY IN", "TRY OUT", "LATIHAN"])
                foto = st.file_uploader("8. Upload Foto Kegiatan (Kamera/Galeri)", type=["jpg", "png", "jpeg"])
                
            st.markdown("---")
            st.subheader("7. Isi Laporan")
            
            # Struktur Laporan Dinamis
            sections = {
                "1. Performa Fisik & Kebugaran": [
                    "A. Capaian parameter fisik (vs benchmark target)",
                    "B. Grafik performa puncak (peaking)",
                    "C. Tingkat pemulihan fisik (recovery)",
                    "D. Keluhan cedera lama / indikasi cedera baru"
                ],
                "2. Kesiapan Taktis & Strategi": [
                    "A. Pemetaan kekuatan calon lawan",
                    "B. Kemampuan mengikuti instruksi teknis",
                    "C. Hasil try-out / sparing"
                ],
                "3. Mental, Psikologis & Kesiapan Mental": [
                    "A. Tingkat kecemasan & pengendalian stres",
                    "B. Fokus, motivasi, dan self-confidence",
                    "C. Rutinitas mental khusus",
                    "D. Koordinasi dengan psikolog olahraga"
                ],
                "4. Nutrisi, Berat Badan & Gaya Hidup": [
                    "A. Progres penyesuaian berat badan",
                    "B. Asupan nutrisi dan suplemen",
                    "C. Status hidrasi",
                    "D. Kualitas dan kecukupan tidur"
                ],
                "5. Medis, Bebas Doping & Logistik": [
                    "A. Status rekam medis & tim medis",
                    "B. Keamanan obat / suplemen (bebas doping)",
                    "C. Kesiapan perlengkapan tanding",
                    "D. Kendala non-teknis"
                ]
            }
            
            laporan_data = {}
            for sec, sub_points in sections.items():
                with st.expander(sec):
                    laporan_data[sec] = {}
                    for point in sub_points:
                        laporan_data[sec][point] = st.text_area(point)
                        
            submit_btn = st.form_submit_button("Buat & Download Laporan Word")
            
            if submit_btn:
                pengurus_list = [p.strip() for p in pengurus_input.split(',')]
                if len(pengurus_list) > 5 and pengurus_list[0] != '':
                    st.error("Jumlah pengurus maksimal 5 orang!")
                else:
                    # Siapkan data untuk Word
                    data_rekap = {
                        "staff": staff,
                        "pengurus": pengurus_input,
                        "cabor": cabor,
                        "tanggal": tanggal.strftime("%d %B %Y"),
                        "tempat": tempat,
                        "kegiatan": kegiatan,
                        "laporan": laporan_data
                    }
                    
                    realtime_stamp = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
                    
                    try:
                        doc_io = generate_word_report(data_rekap, foto, realtime_stamp)
                        st.success("Laporan berhasil disusun!")
                        
                        st.download_button(
                            label="📥 Download Laporan (Word Document)",
                            data=doc_io,
                            file_name=f"Laporan_Monitoring_{cabor}_{tanggal}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"Terjadi kesalahan saat membuat dokumen: {e}")

    elif choice == "Tambah User" and st.session_state["role"] == "admin":
        st.title("Manajemen User")
        st.write("Hanya Admin yang dapat melihat dan menggunakan fitur ini.")
        with st.form("add_user_form"):
            new_user = st.text_input("Username Baru")
            new_pass = st.text_input("Password Baru", type='password')
            role = st.selectbox("Role", ["user", "admin"])
            
            if st.form_submit_button("Tambah"):
                add_userdata(new_user, make_hashes(new_pass), role)
                st.success(f"Berhasil menambahkan user: {new_user}")

    elif choice == "Logout":
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""
        st.rerun()

if __name__ == '__main__':
    main()
