import streamlit as st
from supabase import create_client, Client
import bcrypt
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import urllib.parse
import uuid

# =========================================================
# KONFIGURASI HALAMAN (HARUS DI BARIS PALING ATAS)
# =========================================================
st.set_page_config(
    page_title="Monitoring Binpres KONI",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CUSTOM CSS
# =========================================================
st.markdown("""
    <style>
    .main-header {
        font-size: 38px;
        font-weight: 800;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: -10px;
    }
    .sub-header {
        font-size: 18px;
        font-weight: 400;
        color: #64748B;
        text-align: center;
        margin-bottom: 30px;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        transition: 0.3s;
    }
    .stButton>button:hover {
        transform: scale(1.02);
    }
    .divider {
        height: 2px;
        background: linear-gradient(90deg, #1E3A8A 0%, #3B82F6 100%);
        margin: 20px 0;
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================
# KONFIGURASI SUPABASE
# =========================================================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
BUCKET_NAME = "laporan-monev"

@st.cache_resource
def get_supabase() -> Client:
    if not SUPABASE_URL or not SUPABASE_KEY:
        st.error("❌ Supabase belum dikonfigurasi di Streamlit Secrets.")
        st.stop()
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = get_supabase()

# =========================================================
# WAKTU INDONESIA
# =========================================================
def get_current_time_id():
    now = datetime.datetime.now()
    hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", 
             "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{hari[now.weekday()]}, {now.day} {bulan[now.month - 1]} {now.year} - {now.strftime('%H:%M')} WIB"

# =========================================================
# FUNGSI PASSWORD & DATABASE
# =========================================================
def hash_password(password):
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

def check_password(password, stored_password):
    try:
        return bcrypt.checkpw(password.encode("utf-8"), stored_password.encode("utf-8"))
    except Exception:
        return False

def init_db():
    try:
        result = supabase.table("users").select("username").eq("username", "admin").execute()
        if not result.data:
            supabase.table("users").insert({
                "username": "admin",
                "password": hash_password("admin123"),
                "role": "admin"
            }).execute()
    except Exception as e:
        # Jika error karena RLS, biarkan lewat (karena akun bisa dibuat manual / RLS disable)
        pass

def delete_old_reports():
    try:
        batas = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=30)
        result = supabase.table("reports").select("id, file_path").lt("submit_time", batas.isoformat()).execute()
        old_reports = result.data or []
        for report in old_reports:
            file_path = report.get("file_path")
            if file_path:
                try:
                    supabase.storage.from_(BUCKET_NAME).remove([file_path])
                except:
                    pass
            try:
                supabase.table("reports").delete().eq("id", report["id"]).execute()
            except:
                pass
    except:
        pass

init_db()
delete_old_reports()

def authenticate(username, password):
    try:
        result = supabase.table("users").select("password, role").eq("username", username).execute()
        if not result.data: return False, None
        user = result.data[0]
        if check_password(password, user["password"]):
            return True, user["role"]
    except Exception as e:
        st.error(f"Terjadi kesalahan saat login: {e}")
    return False, None

# =========================================================
# GENERATE WORD REPORT (FORMAT PDF E-MONITORING)
# =========================================================
def generate_word_report(
    cabor, tanggal, tempat, nama_program, fokus, jml_atlet, 
    pelatih, instansi, periode, target, realisasi, status, 
    deskripsi, kendala, mitigasi, admin_nama, admin_jabatan, fotos
):
    doc = Document()

    # Default Font Setting
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # Header Title
    head = doc.add_heading("LAPORAN E-MONITORING OLAHRAGA", level=1)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in head.runs:
        run.font.color.rgb = None  # Remove blue default color
        run.bold = True
    
    subhead = doc.add_paragraph("Sistem Pemantauan Program Pembinaan & Performa Atlet")
    subhead.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # SECTION A: DATA UMUM
    doc.add_heading("A. DATA UMUM PROGRAM", level=2)
    table_a = doc.add_table(rows=8, cols=3)
    table_a.autofit = True
    data_a = [
        ("Nama Program", nama_program),
        ("Cabang Olahraga", cabor),
        ("Fokus Pembinaan", fokus),
        ("Lokasi Pemusatan", tempat),
        ("Jumlah Atlet Aktif", jml_atlet),
        ("Pelatih Kepala", pelatih),
        ("Instansi Pengawas", instansi),
        ("Periode Laporan", periode)
    ]
    for i, (label, val) in enumerate(data_a):
        cells = table_a.rows[i].cells
        cells[0].text = label
        cells[1].text = ":"
        cells[2].text = val
        cells[0].paragraphs[0].runs[0].bold = True
        cells[0].width = Inches(1.8)
        cells[1].width = Inches(0.2)
        cells[2].width = Inches(4.0)

    doc.add_paragraph() # Spacing

    # SECTION B: PROGRESS
    doc.add_heading("B. INDIKATOR KETERCAPAIAN LATIHAN (PROGRESS)", level=2)
    table_b = doc.add_table(rows=2, cols=3, style='Table Grid')
    headers_b = ["Target Kondisi Fisik", "Realisasi Rata-rata Atlet", "Status Evaluasi"]
    for i, header in enumerate(headers_b):
        p = table_b.cell(0, i).paragraphs[0]
        p.add_run(header).bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    vals_b = [target, realisasi, status]
    for i, val in enumerate(vals_b):
        p = table_b.cell(1, i).paragraphs[0]
        p.add_run(val)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph() # Spacing

    # SECTION C: DESKRIPSI
    doc.add_heading("C. DESKRIPSI RINCI PELAKSANAAN PROGRAM", level=2)
    p_desc = doc.add_paragraph(deskripsi)
    p_desc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # SECTION D: KENDALA DAN MITIGASI
    doc.add_heading("D. KENDALA LAPANGAN DAN MITIGASI", level=2)
    table_d = doc.add_table(rows=2, cols=2, style='Table Grid')
    
    # Headers
    h_kendala = table_d.cell(0, 0).paragraphs[0]
    h_kendala.add_run("Identifikasi Kendala").bold = True
    h_kendala.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    h_mitigasi = table_d.cell(0, 1).paragraphs[0]
    h_mitigasi.add_run("Mitigasi & Rencana Tindak Lanjut").bold = True
    h_mitigasi.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Data
    c_kendala = table_d.cell(1, 0).paragraphs[0]
    c_kendala.add_run(kendala)
    c_kendala.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    c_mitigasi = table_d.cell(1, 1).paragraphs[0]
    c_mitigasi.add_run(mitigasi)
    c_mitigasi.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph() # Spacing

    # TANDA TANGAN (Hanya Admin / 1 Tanda Tangan)
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run(f"Dibuat Oleh,\n{admin_jabatan}\n\n\n\n").bold = True
    run_nama = p_sig.add_run(admin_nama)
    run_nama.bold = True
    run_nama.underline = True

    # ==========================
    # HALAMAN 2: LAMPIRAN FOTO
    # ==========================
    if fotos:
        doc.add_page_break()
        head_doc = doc.add_heading("LAMPIRAN FOTO DOKUMENTASI", level=1)
        head_doc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in head_doc.runs:
            run.font.color.rgb = None
            run.bold = True
            
        doc.add_paragraph("Bukti Visual Pelaksanaan Program Pembinaan Olahraga (e-Monitoring)\n").alignment = WD_ALIGN_PARAGRAPH.CENTER

        table_foto = doc.add_table(rows=0, cols=2)
        table_foto.autofit = False
        
        row_cells = None
        for idx, foto in enumerate(fotos):
            if idx % 2 == 0:
                row_cells = table_foto.add_row().cells
            
            cell = row_cells[idx % 2]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            try:
                image_stream = io.BytesIO(foto.getvalue())
                p.add_run().add_picture(image_stream, width=Inches(2.8))
            except Exception as e:
                p.add_run(f"(Gagal memuat gambar: {e})")
            
            p2 = cell.add_paragraph(f"FOTO {idx+1}")
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p2.runs[0].bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =========================================================
# UPLOAD, DOWNLOAD, DELETE FILE DARI STORAGE
# =========================================================
def upload_report_file(file_bytes, file_name):
    unique_folder = datetime.datetime.now().strftime("%Y/%m")
    unique_id = uuid.uuid4().hex[:12]
    storage_path = f"{unique_folder}/{unique_id}_{file_name}"
    try:
        supabase.storage.from_(BUCKET_NAME).upload(
            storage_path, file_bytes,
            {"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "upsert": "false"}
        )
        return storage_path
    except Exception as e:
        raise Exception(f"Gagal upload file ke Storage: {e}")

# =========================================================
# SESSION STATE
# =========================================================
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "username" not in st.session_state: st.session_state["username"] = ""
if "role" not in st.session_state: st.session_state["role"] = ""
if "report_generated" not in st.session_state: st.session_state["report_generated"] = False

# =========================================================
# HALAMAN LOGIN
# =========================================================
if not st.session_state["logged_in"]:
    st.markdown("<div class='main-header'>🏆 E-MONEV CABOR</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Binpres KONI Kabupaten Tangerang</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1.5, 2, 1.5])
    with col2:
        with st.container(border=True):
            st.markdown("#### 🔐 Silakan Masuk")
            with st.form("login_form"):
                username_input = st.text_input("👤 Username")
                password_input = st.text_input("🔑 Password", type="password")
                submit_btn = st.form_submit_button("Masuk Sistem", use_container_width=True)
                
                if submit_btn:
                    if not username_input or not password_input:
                        st.warning("⚠️ Username dan password harus diisi.")
                    else:
                        is_auth, role = authenticate(username_input, password_input)
                        if is_auth:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = username_input
                            st.session_state["role"] = role
                            st.rerun()
                        else:
                            st.error("🚨 Username atau password salah!")

# =========================================================
# HALAMAN SETELAH LOGIN (DASHBOARD)
# =========================================================
else:
    # SIDEBAR
    st.sidebar.markdown("### 🏆 PANEL MONEV")
    st.sidebar.caption("Binpres KONI Kab. Tangerang")
    st.sidebar.markdown(f"**🕒 Waktu Sistem:**\n*{get_current_time_id()}*")
    st.sidebar.markdown("---")
    st.sidebar.info(f"👤 **Login:** {st.session_state['username'].upper()}\n\n🛡️ **Role:** {st.session_state['role'].upper()}")
    
    if st.session_state["role"] == "admin":
        menu = ["📝 Form Laporan e-Monitoring", "📅 Kelola Jadwal", "👥 Kelola User", "📂 Arsip Laporan"]
        choice = st.sidebar.radio("📌 Navigasi Admin:", menu)
    else:
        choice = "📝 Form Laporan e-Monitoring"
        st.sidebar.success("✅ Silakan isi form laporan di panel kanan.")
        
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Keluar (Logout)", use_container_width=True, type="secondary"):
        st.session_state.clear()
        st.rerun()

    # =====================================================
    # MENU 1: FORM LAPORAN E-MONITORING
    # =====================================================
    if choice == "📝 Form Laporan e-Monitoring":
        st.markdown("### 📝 Form Laporan e-Monitoring Olahraga")
        st.markdown(f"**Tanggal Hari Ini:** {get_current_time_id()}")
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

        try:
            schedule_response = supabase.table("schedules").select("id, cabor, tanggal, tempat").order("id", desc=True).execute()
            schedules_data = schedule_response.data or []
        except Exception:
            schedules_data = []

        if not schedules_data:
            st.warning("⚠️ Belum ada jadwal monitoring. Harap tambahkan di menu Kelola Jadwal.")
        else:
            schedule_options = {f"{s['cabor']} | {s['tanggal']} | {s['tempat']}": s for s in schedules_data}
            selected_label = st.selectbox("📌 1. Pilih Jadwal Terdaftar", list(schedule_options.keys()))
            selected_schedule = schedule_options[selected_label]
            
            val_cabor = selected_schedule["cabor"]
            val_tanggal = selected_schedule["tanggal"]
            val_tempat = selected_schedule["tempat"]

            with st.container(border=True):
                st.markdown("#### 📋 2. Formulir Data Program & Evaluasi")
                
                colA, colB = st.columns(2)
                with colA:
                    nama_program = st.text_input("Nama Program", "Pemusatan Latihan Daerah (Pelatda) Utama")
                    fokus = st.text_input("Fokus Pembinaan", "Persiapan Menuju Pekan Olahraga Provinsi (Porprov)")
                    jml_atlet = st.text_input("Jumlah Atlet Aktif", "12 Atlet (7 Putra, 5 Putri)")
                    pelatih = st.text_input("Pelatih Kepala", "")
                with colB:
                    instansi = st.text_input("Instansi Pengawas", "Binpres KONI Kab. Tangerang")
                    periode = st.text_input("Periode Laporan", "September 2026")
                    target_fisik = st.text_input("Target Kondisi Fisik / VO2Max", "90.00%")
                    realisasi = st.text_input("Realisasi Rata-rata Atlet", "87.50%")
                
                status_eval = st.selectbox("Status Evaluasi Latihan", ["Sangat Baik", "Tercapai", "Perlu Peningkatan", "Buruk"])

                deskripsi = st.text_area(
                    "Deskripsi Rinci Pelaksanaan Program", 
                    "Berdasarkan data pemantauan minggu ini, program latihan berjalan sesuai kurikulum...\n"
                    "1. Latihan Fisik (Strength & Conditioning): ...\n"
                    "2. Latihan Teknik: ...\n"
                    "3. Pemulihan (Recovery) & Medis: ...", height=120
                )

                colC, colD = st.columns(2)
                with colC:
                    kendala = st.text_area("Identifikasi Kendala Lapangan", "1. ...\n2. ...")
                with colD:
                    mitigasi = st.text_area("Mitigasi & Rencana Tindak Lanjut", "1. ...\n2. ...")

                st.markdown("#### ✍️ 3. Pengaturan Tanda Tangan Laporan")
                col_sig1, col_sig2 = st.columns(2)
                with col_sig1:
                    admin_nama = st.text_input("Nama Penandatangan", "Dr. Haryanto Saputra, M.Si.")
                with col_sig2:
                    admin_jabatan = st.text_input("Jabatan / Peran", "Ketua Satlak Pembinaan Prestasi")

                st.markdown("**📸 4. Upload Foto Dokumentasi (Min. 1 Foto)**")
                fotos = st.file_uploader(
                    "Foto akan di-layout menjadi kotak berdampingan di Halaman Lampiran.",
                    type=["png", "jpg", "jpeg"], accept_multiple_files=True
                )

                submit_laporan = st.button("📄 Generate & Simpan Laporan", use_container_width=True, type="primary")

            # VALIDASI & GENERATE LAPORAN
            if submit_laporan:
                if not admin_nama.strip(): st.error("⚠️ Nama Penandatangan tidak boleh kosong!")
                elif not fotos or len(fotos) < 1: st.error("🚨 Minimal unggah 1 foto dokumentasi.")
                else:
                    try:
                        with st.spinner("⏳ Menyusun dokumen laporan..."):
                            word_file = generate_word_report(
                                val_cabor, val_tanggal, val_tempat, nama_program, fokus, 
                                jml_atlet, pelatih, instansi, periode, target_fisik, realisasi, 
                                status_eval, deskripsi, kendala, mitigasi, admin_nama, admin_jabatan, fotos
                            )
                            
                            safe_cabor = val_cabor.replace("/", "_").replace("\\", "_").replace(" ", "_")
                            safe_date = val_tanggal.replace(" s/d ", "_").replace("-", "").replace("/", "")
                            file_name_doc = f"Monev_{safe_cabor}_{safe_date}.docx"
                            file_bytes = word_file.getvalue()
                            
                            file_path = upload_report_file(file_bytes, file_name_doc)
                            
                            supabase.table("reports").insert({
                                "cabor": val_cabor,
                                "tanggal_kegiatan": val_tanggal,
                                "file_name": file_name_doc,
                                "file_path": file_path,
                                "submitted_by": st.session_state["username"]
                            }).execute()
                            
                            st.session_state["report_generated"] = True
                            st.session_state["word_file"] = file_bytes
                            st.session_state["file_name_doc"] = file_name_doc

                            pesan = f"Halo Admin, Laporan e-Monitoring *{val_cabor}* telah di-submit ke sistem."
                            st.session_state["wa_link"] = f"https://wa.me/6285691860578?text={urllib.parse.quote(pesan)}"

                        st.success("🎉 Laporan berhasil disimpan ke database!")
                    except Exception as e:
                        st.error("❌ Gagal menyimpan laporan.")
                        st.code(str(e))

            if st.session_state.get("report_generated", False):
                colDL1, colDL2 = st.columns(2)
                with colDL1:
                    st.download_button(
                        label="📥 Unduh File Ms. Word (.docx)", 
                        data=st.session_state["word_file"],
                        file_name=st.session_state["file_name_doc"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with colDL2:
                    st.link_button("📲 Notifikasi Admin via WhatsApp", st.session_state["wa_link"], use_container_width=True)


    # =====================================================
    # MENU 2: KELOLA JADWAL
    # =====================================================
    elif choice == "📅 Kelola Jadwal":
        st.markdown("### 📅 Kelola Jadwal Monitoring")
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        
        daftar_cabor = [
            "ANGGAR (IKASI)", "AERO SPORT (FASI)", "ARUNG JERAM (FAJI)", "ATLETIK (PASI)", 
            "ANGKAT BESI (PABSI)", "ANGKAT BERAT (PABERSI)", "BINARAGA FITNESS (PBFI)", 
            "BILIAR (POBSI)", "BALAP SEPEDA (ISSI)", "BOLA BASKET (PERBASI)", 
            "BOLA SUNDUL (PERBOSI)", "BOLA VOLI (PBVSI)", "BOWLING (PBI)", 
            "BRIDGE (GABSI)", "BULU TANGKIS (PBSI)", "BASEBALL & SOFTBALL (PERBASASI)", 
            "BOLA TANGAN (ABTI)", "CATUR (PERCASI)", "CRICKET (PCI)", "DAYUNG (PODSI)", 
            "DRUM BAND (PDBI)", "GOLF (PGI)", "GULAT (PGSI)", "GATEBALL (PERGATSI)", 
            "HOCKEY (FHI)", "JUDO (PJSI)", "KEMPO (PERKEMI)", "KARATE (FORKI)", 
            "LAYAR (PORLASI)", "MENEMBAK (PERBAKIN)", "MUAY THAI (MI)", "MOTOR (IMI)", 
            "PANAHAN (PERPANI)", "PANJAT TEBING (FPTI)", "PENCAK SILAT (IPSI)", 
            "PETANQUE (POPI)", "RENANG (PRSI)", "RUGBY (PRUI)", "SENAM (PERSANI)", 
            "SEPAK BOLA (Askab-PSSI)", "SEPAK TAKRAW (PSTI)", "SEPATU RODA (PORSEROSI)", 
            "SQUASH (PSI)", "TAEKWONDO (TI)", "TARUNG DERAJAT (KODRAT)", 
            "TENIS LAPANGAN (PELTI)", "TENIS MEJA (PTMSI)", "TINJU (PERTINA)", 
            "WUSHU (WI)", "WOODBALL (IwBA)", "KICKBOXING (KBI)", "E. SPORT", 
            "FLOOR BALL", "MMA", "SELAM", "BARONGSAI (FOBI)", "JUJITSU (PBJI)", 
            "KURASH", "PICKLE BALL", "BAPOPSI", "PERWOSI", "SIWO"
        ]

        with st.form("tambah_jadwal_form"):
            st.subheader("➕ Tambah Jadwal Baru")
            c_cabor = st.selectbox("Cabang Olahraga", daftar_cabor)
            c_tanggal = st.date_input("Tanggal Kegiatan", value=[])
            c_tempat = st.text_input("Lokasi / Tempat", placeholder="Contoh: Stadion Utama")
            
            submit_jadwal = st.form_submit_button("Simpan Jadwal", type="primary")
            if submit_jadwal:
                if c_cabor and c_tempat and len(c_tanggal) > 0:
                    if len(c_tanggal) == 1:
                        tanggal_str = c_tanggal[0].strftime("%d %b %Y")
                    else:
                        tanggal_str = f"{c_tanggal[0].strftime('%d %b %Y')} s/d {c_tanggal[1].strftime('%d %b %Y')}"

                    try:
                        supabase.table("schedules").insert({
                            "cabor": c_cabor,
                            "tanggal": tanggal_str,
                            "tempat": c_tempat
                        }).execute()
                        st.success(f"✅ Jadwal {c_cabor} berhasil ditambahkan!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menambah jadwal: {e}")
                else:
                    st.warning("⚠️ Cabang Olahraga, Tanggal, dan Tempat harus diisi lengkap!")

        st.markdown("#### 📋 Daftar Jadwal Saat Ini")
        try:
            jadwal_data = supabase.table("schedules").select("*").order("id", desc=True).execute().data
            if jadwal_data:
                df_jadwal = pd.DataFrame(jadwal_data)
                st.dataframe(df_jadwal[["cabor", "tanggal", "tempat"]], use_container_width=True)
            else:
                st.info("Belum ada jadwal yang terdaftar.")
        except:
            st.info("Tabel 'schedules' belum tersedia atau kosong.")


    # =====================================================
    # MENU 3: KELOLA USER
    # =====================================================
    elif choice == "👥 Kelola User":
        st.markdown("### 👥 Manajemen Pengguna")
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        
        with st.form("tambah_user_form"):
            st.subheader("➕ Tambah Akun Baru")
            u_name = st.text_input("Username Baru")
            u_pass = st.text_input("Password", type="password")
            u_role = st.selectbox("Role (Hak Akses)", ["user", "admin"])
            
            submit_user = st.form_submit_button("Buat Akun", type="primary")
            if submit_user:
                if u_name and u_pass:
                    try:
                        supabase.table("users").insert({
                            "username": u_name.lower(),
                            "password": hash_password(u_pass),
                            "role": u_role
                        }).execute()
                        st.success(f"✅ Akun {u_name} berhasil dibuat!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal membuat akun: {e}")
                else:
                    st.warning("⚠️ Username dan Password tidak boleh kosong!")
                    
        st.markdown("#### 📋 Daftar Akun")
        try:
            users_data = supabase.table("users").select("username, role").execute().data
            if users_data:
                st.dataframe(pd.DataFrame(users_data), use_container_width=True)
        except:
            st.info("Tidak dapat memuat data user.")


    # =====================================================
    # MENU 4: ARSIP LAPORAN
    # =====================================================
    elif choice == "📂 Arsip Laporan":
        st.markdown("### 📂 Arsip Laporan Tersimpan")
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        
        try:
            laporan_data = supabase.table("reports").select("*").order("submit_time", desc=True).execute().data
            if laporan_data:
                for rep in laporan_data:
                    with st.expander(f"📄 {rep['cabor']} - {rep['tanggal_kegiatan']}"):
                        st.write(f"**Disubmit oleh:** {rep['submitted_by']}")
                        st.write(f"**Waktu Arsip:** {rep['submit_time']}")
                        st.write(f"**Nama File:** {rep.get('file_name', 'Tidak diketahui')}")
            else:
                st.info("Belum ada laporan yang tersimpan di sistem.")
        except:
            st.info("Tabel 'reports' belum tersedia atau kosong.")
