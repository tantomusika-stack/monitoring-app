import streamlit as st
from supabase import create_client, Client
import bcrypt
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import urllib.parse
import uuid


# =========================================================
# KONFIGURASI HALAMAN
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

    .info-box {
        background-color: #F8FAFC;
        border-left: 5px solid #3B82F6;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
    }

    .divider {
        height: 2px;
        background: linear-gradient(
            90deg,
            #1E3A8A 0%,
            #3B82F6 100%
        );
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
        st.error(
            "❌ Supabase belum dikonfigurasi. "
            "Silakan masukkan SUPABASE_URL dan SUPABASE_KEY "
            "di Streamlit Secrets."
        )
        st.stop()

    return create_client(
        SUPABASE_URL,
        SUPABASE_KEY
    )


supabase = get_supabase()


# =========================================================
# WAKTU INDONESIA
# =========================================================

def get_current_time_id():

    now = datetime.datetime.now()

    hari = [
        "Senin",
        "Selasa",
        "Rabu",
        "Kamis",
        "Jumat",
        "Sabtu",
        "Minggu"
    ]

    bulan = [
        "Januari",
        "Februari",
        "Maret",
        "April",
        "Mei",
        "Juni",
        "Juli",
        "Agustus",
        "September",
        "Oktober",
        "November",
        "Desember"
    ]

    return (
        f"{hari[now.weekday()]}, "
        f"{now.day} "
        f"{bulan[now.month - 1]} "
        f"{now.year} - "
        f"{now.strftime('%H:%M')} WIB"
    )


# =========================================================
# FUNGSI PASSWORD
# =========================================================

def hash_password(password):

    hashed = bcrypt.hashpw(
        password.encode("utf-8"),
        bcrypt.gensalt()
    )

    return hashed.decode("utf-8")


def check_password(password, stored_password):

    try:

        return bcrypt.checkpw(
            password.encode("utf-8"),
            stored_password.encode("utf-8")
        )

    except Exception:

        return False


# =========================================================
# INISIALISASI DATABASE
# =========================================================

def init_db():

    try:

        # Cek admin
        result = (
            supabase
            .table("users")
            .select("username")
            .eq("username", "admin")
            .execute()
        )

        if not result.data:

            supabase.table("users").insert({
                "username": "admin",
                "password": hash_password("admin123"),
                "role": "admin"
            }).execute()

    except Exception as e:

        st.error(
            "❌ Gagal menghubungkan database Supabase."
        )

        st.code(str(e))

        st.stop()


# =========================================================
# HAPUS LAPORAN LEBIH DARI 30 HARI
# =========================================================

def delete_old_reports():

    try:

        batas = (
            datetime.datetime.now(datetime.timezone.utc)
            - datetime.timedelta(days=30)
        )

        result = (
            supabase
            .table("reports")
            .select("id, file_path")
            .lt(
                "submit_time",
                batas.isoformat()
            )
            .execute()
        )

        old_reports = result.data or []

        for report in old_reports:

            file_path = report.get("file_path")

            if file_path:

                try:

                    supabase.storage \
                        .from_(BUCKET_NAME) \
                        .remove([file_path])

                except Exception:

                    pass

            try:

                (
                    supabase
                    .table("reports")
                    .delete()
                    .eq("id", report["id"])
                    .execute()
                )

            except Exception:

                pass

    except Exception:

        pass


# Jalankan database
init_db()

# Hapus otomatis laporan >30 hari
delete_old_reports()


# =========================================================
# AUTHENTICATION
# =========================================================

def authenticate(username, password):

    try:

        result = (
            supabase
            .table("users")
            .select("password, role")
            .eq("username", username)
            .execute()
        )

        if not result.data:

            return False, None

        user = result.data[0]

        stored_password = user["password"]
        role = user["role"]

        if check_password(
            password,
            stored_password
        ):

            return True, role

    except Exception as e:

        st.error(
            f"Terjadi kesalahan saat login: {e}"
        )

    return False, None


# =========================================================
# GENERATE WORD REPORT
# =========================================================

def generate_word_report(
    petugas_text,
    cabor,
    tanggal,
    tempat,
    catatan,
    fotos
):

    doc = Document()

    # Header
    head = doc.add_heading(
        "LAPORAN MONITORING CABANG OLAHRAGA",
        0
    )

    head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Informasi jadwal
    doc.add_paragraph(
        f"Cabang Olahraga\t: {cabor}"
    )

    doc.add_paragraph(
        f"Tanggal\t\t: {tanggal}"
    )

    doc.add_paragraph(
        f"Tempat\t\t: {tempat}\n"
    )

    # Petugas
    doc.add_heading(
        "Daftar Petugas:",
        level=3
    )

    petugas_list = [
        p.strip()
        for p in petugas_text.split("\n")
        if p.strip()
    ]

    for i, p in enumerate(
        petugas_list,
        1
    ):

        doc.add_paragraph(
            f"{i}. {p}"
        )

    # Catatan
    doc.add_heading(
        "Catatan Evaluasi / Hasil Monitoring:",
        level=3
    )

    doc.add_paragraph(catatan)

    # Foto
    if fotos:

        doc.add_page_break()

        head_doc = doc.add_heading(
            "Lampiran Foto Dokumentasi",
            level=2
        )

        head_doc.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph(
            f"Cabor: {cabor} | Tanggal: {tanggal}\n"
        )

        table = doc.add_table(
            rows=0,
            cols=2
        )

        table.autofit = False

        row_cells = None

        for idx, foto in enumerate(fotos):

            if idx % 2 == 0:

                row_cells = table.add_row().cells

            cell = row_cells[idx % 2]

            p = cell.paragraphs[0]

            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = p.add_run()

            try:

                image_stream = io.BytesIO(
                    foto.getvalue()
                )

                run.add_picture(
                    image_stream,
                    width=Inches(2.8)
                )

            except Exception as e:

                run.add_text(
                    f"(Gagal memuat gambar: {e})"
                )

    buffer = io.BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer


# =========================================================
# UPLOAD FILE KE SUPABASE STORAGE
# =========================================================

def upload_report_file(
    file_bytes,
    file_name
):

    unique_folder = datetime.datetime.now().strftime(
        "%Y/%m"
    )

    unique_id = uuid.uuid4().hex[:12]

    storage_path = (
        f"{unique_folder}/"
        f"{unique_id}_{file_name}"
    )

    try:

        supabase.storage \
            .from_(BUCKET_NAME) \
            .upload(
                storage_path,
                file_bytes,
                {
                    "content-type":
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document",
                    "upsert": "false"
                }
            )

        return storage_path

    except Exception as e:

        raise Exception(
            f"Gagal upload file ke Storage: {e}"
        )


# =========================================================
# DOWNLOAD FILE DARI STORAGE
# =========================================================

def download_report_file(file_path):

    try:

        file_bytes = (
            supabase
            .storage
            .from_(BUCKET_NAME)
            .download(file_path)
        )

        return file_bytes

    except Exception as e:

        raise Exception(
            f"Gagal mengambil file: {e}"
        )


# =========================================================
# DELETE FILE DARI STORAGE
# =========================================================

def delete_report_file(file_path):

    try:

        if file_path:

            (
                supabase
                .storage
                .from_(BUCKET_NAME)
                .remove([file_path])
            )

    except Exception:

        pass


# =========================================================
# SESSION STATE
# =========================================================

if "logged_in" not in st.session_state:

    st.session_state["logged_in"] = False

if "username" not in st.session_state:

    st.session_state["username"] = ""

if "role" not in st.session_state:

    st.session_state["role"] = ""

if "report_generated" not in st.session_state:

    st.session_state["report_generated"] = False


# =========================================================
# HALAMAN LOGIN
# =========================================================

if not st.session_state["logged_in"]:

    st.markdown(
        "<div class='main-header'>🏆 E-MONEV CABOR</div>",
        unsafe_allow_html=True
    )

    st.markdown(
        "<div class='sub-header'>"
        "Binpres KONI Kabupaten Tangerang"
        "</div>",
        unsafe_allow_html=True
    )

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(
        [1.5, 2, 1.5]
    )

    with col2:

        with st.container(border=True):

            st.markdown(
                "#### 🔐 Silakan Masuk"
            )

            with st.form("login_form"):

                username_input = st.text_input(
                    "👤 Username"
                )

                password_input = st.text_input(
                    "🔑 Password",
                    type="password"
                )

                submit_btn = st.form_submit_button(
                    "Masuk Sistem",
                    use_container_width=True
                )

                if submit_btn:

                    if not username_input or not password_input:

                        st.warning(
                            "⚠️ Username dan password harus diisi."
                        )

                    else:

                        is_auth, role = authenticate(
                            username_input,
                            password_input
                        )

                        if is_auth:

                            st.session_state[
                                "logged_in"
                            ] = True

                            st.session_state[
                                "username"
                            ] = username_input

                            st.session_state[
                                "role"
                            ] = role

                            st.rerun()

                        else:

                            st.error(
                                "🚨 Username atau password salah!"
                            )


# =========================================================
# HALAMAN SETELAH LOGIN
# =========================================================

else:

    # =====================================================
    # SIDEBAR
    # =====================================================

    st.sidebar.markdown(
        "### 🏆 PANEL MONEV"
    )

    st.sidebar.caption(
        "Binpres KONI Kab. Tangerang"
    )

    st.sidebar.markdown(
        f"**🕒 Waktu Sistem:**\n"
        f"*{get_current_time_id()}*"
    )

    st.sidebar.markdown("---")

    st.sidebar.info(
        f"👤 **Login:** "
        f"{st.session_state['username'].upper()}\n\n"
        f"🛡️ **Role:** "
        f"{st.session_state['role'].upper()}"
    )

    # =====================================================
    # MENU
    # =====================================================

    if st.session_state["role"] == "admin":

        menu = [
            "📅 Kelola Jadwal (Admin)",
            "👥 Kelola User (Admin)",
            "📂 Arsip Laporan (Admin)",
            "📝 Coba Isi Laporan"
        ]

        choice = st.sidebar.radio(
            "📌 Navigasi Admin:",
            menu
        )

    else:

        choice = "📝 Isi Form Laporan"

        st.sidebar.success(
            "✅ Silakan isi form laporan "
            "di panel kanan."
        )

    st.sidebar.markdown("---")

    if st.sidebar.button(
        "🚪 Keluar (Logout)",
        use_container_width=True,
        type="secondary"
    ):

        st.session_state.clear()

        st.rerun()


    # =====================================================
    # FORM LAPORAN
    # =====================================================

    if choice in [
        "📝 Isi Form Laporan",
        "📝 Coba Isi Laporan"
    ]:

        st.markdown(
            "### 📝 Form Laporan Monitoring"
        )

        st.markdown(
            f"**Tanggal Hari Ini:** "
            f"{get_current_time_id()}"
        )

        st.markdown(
            "<div class='divider'></div>",
            unsafe_allow_html=True
        )

        try:

            schedule_response = (
                supabase
                .table("schedules")
                .select(
                    "id, cabor, tanggal, tempat"
                )
                .order(
                    "id",
                    desc=True
                )
                .execute()
            )

            schedules_data = (
                schedule_response.data or []
            )

        except Exception as e:

            st.error(
                f"❌ Gagal mengambil jadwal: {e}"
            )

            schedules_data = []


        if not schedules_data:

            st.warning(
                "⚠️ Belum ada jadwal monitoring "
                "yang tersedia. Harap hubungi Admin."
            )

        else:

            schedule_options = {}

            for s in schedules_data:

                label = (
                    f"{s['cabor']} | "
                    f"{s['tanggal']} | "
                    f"{s['tempat']}"
                )

                schedule_options[label] = s

            selected_label = st.selectbox(
                "📌 1. Pilih Jadwal Monitoring yang Tersedia",
                list(schedule_options.keys())
            )

            selected_schedule = (
                schedule_options[selected_label]
            )

            val_cabor = selected_schedule["cabor"]
            val_tanggal = selected_schedule["tanggal"]
            val_tempat = selected_schedule["tempat"]

            with st.container(border=True):

                st.markdown(
                    "#### 📋 2. Detail Evaluasi & Dokumentasi"
                )

                petugas_text = st.text_area(
                    "👤 Daftar Petugas "
                    "(Tulis 1 nama per baris)",
                    placeholder=(
                        "Contoh:\n"
                        "Budi Santoso\n"
                        "Andi Saputra"
                    ),
                    height=100
                )

                catatan = st.text_area(
                    "✍️ Catatan Evaluasi / "
                    "Hasil Monitoring",
                    height=150
                )

                st.markdown(
                    "**📸 Upload Foto Bukti "
                    "(Bebas 2 s/d 5 Foto)**"
                )

                fotos = st.file_uploader(
                    "Otomatis digabung jadi "
                    "1 halaman rapi di Word.",
                    type=[
                        "png",
                        "jpg",
                        "jpeg"
                    ],
                    accept_multiple_files=True
                )

                submit_laporan = st.button(
                    "📄 Generate & Simpan Laporan",
                    use_container_width=True,
                    type="primary"
                )


            if submit_laporan:

                if not petugas_text.strip():

                    st.error(
                        "⚠️ Harap isi minimal "
                        "1 nama petugas!"
                    )

                elif not catatan.strip():

                    st.error(
                        "⚠️ Catatan evaluasi "
                        "tidak boleh kosong!"
                    )

                elif not fotos or len(fotos) < 2:

                    st.error(
                        "🚨 Minimal unggah "
                        "2 foto dokumentasi."
                    )

                elif len(fotos) > 5:

                    st.error(
                        "🚨 Maksimal 5 foto "
                        "dokumentasi agar muat "
                        "1 halaman."
                    )

                else:

                    try:

                        with st.spinner(
                            "⏳ Menyusun dokumen "
                            "laporan..."
                        ):

                            # Generate Word
                            word_file = generate_word_report(
                                petugas_text,
                                val_cabor,
                                val_tanggal,
                                val_tempat,
                                catatan,
                                fotos
                            )

                            # Nama file
                            safe_cabor = (
                                val_cabor
                                .replace("/", "_")
                                .replace("\\", "_")
                                .replace(" ", "_")
                            )

                            safe_date_name = (
                                val_tanggal
                                .replace(" s/d ", "_")
                                .replace("-", "")
                                .replace("/", "")
                            )

                            file_name_doc = (
                                f"Monev_"
                                f"{safe_cabor}_"
                                f"{safe_date_name}.docx"
                            )

                            # Bytes file
                            file_bytes = (
                                word_file.getvalue()
                            )

                            # Upload ke Supabase Storage
                            file_path = upload_report_file(
                                file_bytes,
                                file_name_doc
                            )

                            # Simpan metadata
                            report_data = {

                                "cabor": val_cabor,

                                "tanggal_kegiatan":
                                    val_tanggal,

                                "file_name":
                                    file_name_doc,

                                "file_path":
                                    file_path,

                                "submitted_by":
                                    st.session_state[
                                        "username"
                                    ]
                            }

                            (
                                supabase
                                .table("reports")
                                .insert(report_data)
                                .execute()
                            )

                            # Session
                            st.session_state[
                                "report_generated"
                            ] = True

                            st.session_state[
                                "word_file"
                            ] = file_bytes

                            st.session_state[
                                "file_name_doc"
                            ] = file_name_doc

                            # WhatsApp
                            wa_number = "6285691860578"

                            pesan = (
                                f"Halo Admin, "
                                f"Laporan Monitoring "
                                f"*{val_cabor}* "
                                f"(Tanggal Kegiatan: "
                                f"{val_tanggal}) "
                                f"telah selesai dibuat "
                                f"dan berhasil masuk "
                                f"ke sistem.\n\n"
                                f"Silakan login ke aplikasi "
                                f"dan buka menu "
                                f"*Arsip Laporan* "
                                f"untuk mengunduh dokumen."
                            )

                            wa_link = (
                                f"https://wa.me/"
                                f"{wa_number}"
                                f"?text="
                                f"{urllib.parse.quote(pesan)}"
                            )

                            st.session_state[
                                "wa_link"
                            ] = wa_link

                        st.success(
                            "🎉 Laporan berhasil "
                            "disimpan permanen "
                            "ke database!"
                        )

                    except Exception as e:

                        st.error(
                            "❌ Gagal menyimpan laporan."
                        )

                        st.code(str(e))


            # =================================================
            # HASIL GENERATE
            # =================================================

            if st.session_state.get(
                "report_generated",
                False
            ):

                st.success(
                    "🎉 **Laporan Berhasil "
                    "Disimpan di Sistem!**"
                )

                colA, colB = st.columns(2)

                with colA:

                    st.download_button(
                        label=(
                            "📥 Unduh Salinan "
                            "untuk Anda"
                        ),

                        data=st.session_state[
                            "word_file"
                        ],

                        file_name=st.session_state[
                            "file_name_doc"
                        ],

                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.wordprocessingml.document"
                        ),

                        use_container_width=True
                    )

                with colB:

                    st.link_button(
                        "📲 Kirim Notifikasi "
                        "via WhatsApp ke Admin",

                        st.session_state[
                            "wa_link"
                        ],

                        use_container_width=True
                    )

                    st.caption(
                        "*(Kirim pesan teks ini agar "
                        "Admin tahu laporan sudah "
                        "siap diunduh)*"
                    )


    # =====================================================
    # ARSIP LAPORAN ADMIN
    # =====================================================

    elif choice == "📂 Arsip Laporan (Admin)":

        st.markdown(
            "### 📂 Arsip Laporan Tersimpan"
        )

        st.markdown(
            "⚠️ *Laporan yang berusia lebih "
            "dari 30 hari akan otomatis "
            "dihapus oleh sistem.*"
        )

        st.markdown(
            "<div class='divider'></div>",
            unsafe_allow_html=True
        )

        try:

            reports_response = (
                supabase
                .table("reports")
                .select(
                    "id, cabor, tanggal_kegiatan, "
                    "submit_time, file_name, "
                    "submitted_by"
                )
                .order(
                    "submit_time",
                    desc=True
                )
                .execute()
            )

            reports = (
                reports_response.data or []
            )

        except Exception as e:

            st.error(
                f"❌ Gagal mengambil arsip: {e}"
            )

            reports = []


        if not reports:

            st.info(
                "Belum ada laporan yang "
                "di-submit dan tersimpan "
                "di sistem saat ini."
            )

        else:

            reports_df = pd.DataFrame(
                reports
            )

            reports_df = reports_df.rename(
                columns={
                    "id": "ID",
                    "cabor": "Cabor",
                    "tanggal_kegiatan":
                        "Tgl Kegiatan",
                    "submit_time":
                        "Waktu Submit",
                    "file_name":
                        "Nama File",
                    "submitted_by":
                        "Petugas"
                }
            )

            display_columns = [
                "ID",
                "Cabor",
                "Tgl Kegiatan",
                "Waktu Submit",
                "Nama File",
                "Petugas"
            ]

            st.dataframe(
                reports_df[
                    display_columns
                ],
                use_container_width=True,
                hide_index=True
            )

            col_dl, col_del = st.columns(2)

            # =============================================
            # DOWNLOAD
            # =============================================

            with col_dl:

                with st.container(
                    border=True
                ):

                    st.markdown(
                        "#### 📥 Unduh Laporan"
                    )

                    report_options = {}

                    for report in reports:

                        label = (
                            f"ID {report['id']} | "
                            f"{report['cabor']} | "
                            f"{report['tanggal_kegiatan']}"
                        )

                        report_options[label] = report

                    selected_report_label = (
                        st.selectbox(
                            "Pilih laporan yang "
                            "ingin diunduh:",
                            list(
                                report_options.keys()
                            )
                        )
                    )

                    selected_report = (
                        report_options[
                            selected_report_label
                        ]
                    )

                    if st.button(
                        "📥 Siapkan File Download",
                        use_container_width=True,
                        type="primary"
                    ):

                        try:

                            with st.spinner(
                                "⏳ Mengambil file..."
                            ):

                                file_bytes = (
                                    download_report_file(
                                        selected_report[
                                            "file_path"
                                        ]
                                    )
                                )

                            st.download_button(
                                label=(
                                    f"⬇️ Unduh "
                                    f"'{selected_report['file_name']}'"
                                ),

                                data=file_bytes,

                                file_name=(
                                    selected_report[
                                        "file_name"
                                    ]
                                ),

                                mime=(
                                    "application/vnd.openxmlformats-"
                                    "officedocument.wordprocessingml.document"
                                ),

                                use_container_width=True
                            )

                        except Exception as e:

                            st.error(
                                f"❌ Gagal mengunduh "
                                f"file: {e}"
                            )


            # =============================================
            # DELETE
            # =============================================

            with col_del:

                with st.container(
                    border=True
                ):

                    st.markdown(
                        "#### 🗑️ Hapus Laporan Manual"
                    )

                    delete_options = {}

                    for report in reports:

                        label = (
                            f"ID {report['id']} | "
                            f"{report['cabor']} | "
                            f"{report['tanggal_kegiatan']}"
                        )

                        delete_options[label] = report

                    delete_label = st.selectbox(
                        "Pilih laporan yang "
                        "akan dihapus:",
                        list(
                            delete_options.keys()
                        ),
                        key="delete_report_select"
                    )

                    selected_delete_report = (
                        delete_options[
                            delete_label
                        ]
                    )

                    confirm_delete = st.checkbox(
                        "Saya yakin ingin "
                        "menghapus laporan ini.",
                        key="confirm_delete_report"
                    )

                    if st.button(
                        "🗑️ Hapus File Ini",
                        use_container_width=True,
                        type="primary"
                    ):

                        if not confirm_delete:

                            st.warning(
                                "⚠️ Centang konfirmasi "
                                "terlebih dahulu."
                            )

                        else:

                            try:

                                report_id = (
                                    selected_delete_report[
                                        "id"
                                    ]
                                )

                                file_path = (
                                    selected_delete_report[
                                        "file_path"
                                    ]
                                )

                                # Hapus storage
                                delete_report_file(
                                    file_path
                                )

                                # Hapus database
                                (
                                    supabase
                                    .table("reports")
                                    .delete()
                                    .eq(
                                        "id",
                                        report_id
                                    )
                                    .execute()
                                )

                                st.success(
                                    f"✅ Laporan ID "
                                    f"{report_id} "
                                    f"berhasil dihapus."
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"❌ Gagal menghapus "
                                    f"laporan: {e}"
                                )


    # =====================================================
    # KELOLA JADWAL
    # =====================================================

    elif choice == "📅 Kelola Jadwal (Admin)":

        st.markdown(
            "### 📅 Kelola Jadwal Monitoring"
        )

        st.markdown(
            f"**Waktu Saat Ini:** "
            f"{get_current_time_id()}"
        )

        st.markdown(
            "<div class='divider'></div>",
            unsafe_allow_html=True
        )

        base_cabor = [

            "ANGGAR (IKASI)",
            "AERO SPORT (FASI)",
            "ARUNG JERAM (FAJI)",
            "ATLETIK (PASI)",
            "ANGKAT BESI (PABSI)",
            "ANGKAT BERAT (PABERSI)",
            "BINARAGA FITNESS (PBFI)",
            "BILIAR (POBSI)",
            "BALAP SEPEDA (ISSI)",
            "BOLA BASKET (PERBASI)",
            "BOLA SUNDUL (PERBOSI)",
            "BOLA VOLI (PBVSI)",
            "BOWLING (PBI)",
            "BRIDGE (GABSI)",
            "BULU TANGKIS (PBSI)",
            "BASEBALL & SOFTBALL (PERBASASI)",
            "BOLA TANGAN (ABTI)",
            "CATUR (PERCASI)",
            "CRICKET (PCI)",
            "DAYUNG (PODSI)",
            "DRUM BAND (PDBI)",
            "GOLF (PGI)",
            "GULAT (PGSI)",
            "GATEBALL (PERGATSI)",
            "HOCKEY (FHI)",
            "JUDO (PJSI)",
            "KEMPO (PERKEMI)",
            "KARATE (FORKI)",
            "LAYAR (PORLASI)",
            "MENEMBAK (PERBAKIN)",
            "MUAY THAI (M I)",
            "MOTOR (I M I)",
            "PANAHAN (PERPANI)",
            "PANJAT TEBING (FPTI)",
            "PENCAK SILAT (IPSI)",
            "PETANQUE (POPI)",
            "RENANG (PRSI)",
            "RUGBY (PRUI)",
            "SENAM (PERSANI)",
            "SEPAK BOLA (Askab-PSSI)",
            "SEPAK TAKRAW (PSTI)",
            "SEPATU RODA (PORSEROSI)",
            "SQUASH (P S I)",
            "TAEKWONDO (T I)",
            "TARUNG DERAJAT (KODRAT)",
            "TENIS LAPANGAN (PELTI)",
            "TENIS MEJA (PTMSI)",
            "TINJU (PERTINA)",
            "WUSHU (W I)",
            "WOODBALL (IwBA)",
            "KICKBOXING (KBI)",
            "E. SPORT",
            "FLOOR BALL",
            "MMA",
            "SELAM",
            "BARONGSAI (FOBI)",
            "JUJITSU (PBJI)",
            "KURASH",
            "PIKCLE BALL",
            "BAPOPSI",
            "PERWOSI",
            "SIWO"
        ]

        try:

            existing_response = (
                supabase
                .table("schedules")
                .select("cabor")
                .execute()
            )

            existing_cabors = [
                row["cabor"]
                for row in (
                    existing_response.data or []
                )
            ]

        except Exception:

            existing_cabors = []


        combined_cabor = sorted(
            list(
                set(
                    base_cabor +
                    existing_cabors
                )
            )
        )

        combined_cabor.append(
            "➕ LAINNYA (Tambah Baru)"
        )

        col_form, col_data = st.columns(
            [1, 1.5]
        )


        # =================================================
        # FORM TAMBAH JADWAL
        # =================================================

        with col_form:

            with st.container(
                border=True
            ):

                st.subheader(
                    "➕ Tambah Jadwal Baru"
                )

                with st.form(
                    "form_jadwal"
                ):

                    selected_cabor_option = (
                        st.selectbox(
                            "Pilih Cabang Olahraga",
                            combined_cabor
                        )
                    )

                    if (
                        selected_cabor_option
                        ==
                        "➕ LAINNYA (Tambah Baru)"
                    ):

                        custom_cabor = st.text_input(
                            "Ketik Nama Cabor Baru",
                            placeholder=(
                                "Cth: PANAHAN (PERPANI)"
                            )
                        )

                    else:

                        custom_cabor = ""


                    new_tanggal = st.date_input(
                        "Tanggal Kegiatan "
                        "(Bisa pilih satu hari "
                        "atau rentang hari)",
                        value=(
                            datetime.date.today(),
                            datetime.date.today()
                        )
                    )

                    new_tempat = st.text_input(
                        "Tempat / Lokasi"
                    )

                    submit_jadwal = (
                        st.form_submit_button(
                            "Simpan Jadwal",
                            use_container_width=True
                        )
                    )


                    if submit_jadwal:

                        final_cabor = (
                            custom_cabor
                            .strip()
                            .upper()
                            if
                            selected_cabor_option
                            ==
                            "➕ LAINNYA (Tambah Baru)"
                            else
                            selected_cabor_option
                        )


                        # Tanggal
                        if isinstance(
                            new_tanggal,
                            tuple
                        ):

                            if len(
                                new_tanggal
                            ) == 2:

                                if (
                                    new_tanggal[0]
                                    ==
                                    new_tanggal[1]
                                ):

                                    final_tanggal = (
                                        new_tanggal[0]
                                        .strftime(
                                            "%d-%m-%Y"
                                        )
                                    )

                                else:

                                    final_tanggal = (
                                        f"{new_tanggal[0].strftime('%d-%m-%Y')}"
                                        f" s/d "
                                        f"{new_tanggal[1].strftime('%d-%m-%Y')}"
                                    )

                            elif len(
                                new_tanggal
                            ) == 1:

                                final_tanggal = (
                                    new_tanggal[0]
                                    .strftime(
                                        "%d-%m-%Y"
                                    )
                                )

                            else:

                                final_tanggal = ""

                        else:

                            final_tanggal = (
                                new_tanggal
                                .strftime(
                                    "%d-%m-%Y"
                                )
                            )


                        if (
                            selected_cabor_option
                            ==
                            "➕ LAINNYA (Tambah Baru)"
                            and not final_cabor
                        ):

                            st.warning(
                                "⚠️ Nama Cabang "
                                "Olahraga baru "
                                "tidak boleh kosong!"
                            )

                        elif not new_tempat:

                            st.warning(
                                "⚠️ Tempat/Lokasi "
                                "tidak boleh kosong!"
                            )

                        elif not final_tanggal:

                            st.warning(
                                "⚠️ Tanggal kegiatan "
                                "tidak boleh kosong!"
                            )

                        else:

                            try:

                                supabase \
                                    .table(
                                        "schedules"
                                    ) \
                                    .insert({
                                        "cabor":
                                            final_cabor,
                                        "tanggal":
                                            final_tanggal,
                                        "tempat":
                                            new_tempat
                                    }) \
                                    .execute()

                                st.success(
                                    f"✅ Jadwal "
                                    f"{final_cabor} "
                                    f"ditambahkan!"
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"❌ Gagal menyimpan "
                                    f"jadwal: {e}"
                                )


        # =================================================
        # DAFTAR JADWAL
        # =================================================

        with col_data:

            with st.container(
                border=True
            ):

                st.subheader(
                    "📋 Daftar Jadwal Aktif"
                )

                try:

                    jadwal_response = (
                        supabase
                        .table("schedules")
                        .select(
                            "id, cabor, tanggal, tempat"
                        )
                        .order(
                            "id",
                            desc=True
                        )
                        .execute()
                    )

                    jadwal_data = (
                        jadwal_response.data or []
                    )

                except Exception as e:

                    st.error(
                        f"❌ Gagal mengambil "
                        f"jadwal: {e}"
                    )

                    jadwal_data = []


                if not jadwal_data:

                    st.info(
                        "Belum ada jadwal "
                        "yang dibuat."
                    )

                else:

                    jadwal_df = pd.DataFrame(
                        jadwal_data
                    )

                    jadwal_df = jadwal_df.rename(
                        columns={
                            "id": "ID",
                            "cabor": "Cabor",
                            "tanggal": "Tanggal",
                            "tempat": "Tempat"
                        }
                    )

                    st.dataframe(
                        jadwal_df[
                            [
                                "ID",
                                "Cabor",
                                "Tanggal",
                                "Tempat"
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True
                    )

                    with st.expander(
                        "🗑️ Hapus Jadwal"
                    ):

                        jadwal_options = {}

                        for row in jadwal_data:

                            label = (
                                f"ID {row['id']} | "
                                f"{row['cabor']} | "
                                f"{row['tanggal']}"
                            )

                            jadwal_options[
                                label
                            ] = row


                        delete_schedule_label = (
                            st.selectbox(
                                "Pilih ID Jadwal "
                                "yang akan dihapus",
                                list(
                                    jadwal_options.keys()
                                )
                            )
                        )

                        selected_schedule_delete = (
                            jadwal_options[
                                delete_schedule_label
                            ]
                        )


                        if st.button(
                            "Hapus Jadwal",
                            type="primary"
                        ):

                            try:

                                (
                                    supabase
                                    .table(
                                        "schedules"
                                    )
                                    .delete()
                                    .eq(
                                        "id",
                                        selected_schedule_delete[
                                            "id"
                                        ]
                                    )
                                    .execute()
                                )

                                st.success(
                                    "✅ Jadwal "
                                    "berhasil dihapus!"
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"❌ Gagal menghapus "
                                    f"jadwal: {e}"
                                )


    # =====================================================
    # KELOLA USER
    # =====================================================

    elif choice == "👥 Kelola User (Admin)":

        st.markdown(
            "### 👥 Manajemen Pengguna"
        )

        st.markdown(
            "<div class='divider'></div>",
            unsafe_allow_html=True
        )

        tab1, tab2 = st.tabs(
            [
                "📋 Daftar Pengguna",
                "➕ Tambah Pengguna Baru"
            ]
        )


        # =================================================
        # DAFTAR USER
        # =================================================

        with tab1:

            try:

                users_response = (
                    supabase
                    .table("users")
                    .select(
                        "username, role"
                    )
                    .order(
                        "username"
                    )
                    .execute()
                )

                users_data = (
                    users_response.data or []
                )

            except Exception as e:

                st.error(
                    f"❌ Gagal mengambil "
                    f"data pengguna: {e}"
                )

                users_data = []


            if users_data:

                users_df = pd.DataFrame(
                    users_data
                )

                users_df = users_df.rename(
                    columns={
                        "username":
                            "Username",
                        "role":
                            "Hak Akses"
                    }
                )

                st.dataframe(
                    users_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "Belum ada pengguna."
                )


            with st.expander(
                "🗑️ Hapus Pengguna",
                expanded=False
            ):

                if users_data:

                    usernames = [
                        u["username"]
                        for u in users_data
                    ]

                    del_user = st.selectbox(
                        "Pilih pengguna "
                        "yang akan dihapus",
                        usernames
                    )

                    if st.button(
                        "Hapus Akun",
                        type="primary"
                    ):

                        if del_user == "admin":

                            st.error(
                                "⚠️ Tidak bisa "
                                "menghapus akun "
                                "admin utama!"
                            )

                        else:

                            try:

                                (
                                    supabase
                                    .table("users")
                                    .delete()
                                    .eq(
                                        "username",
                                        del_user
                                    )
                                    .execute()
                                )

                                st.success(
                                    f"✅ User "
                                    f"**{del_user}** "
                                    f"berhasil dihapus!"
                                )

                                st.rerun()

                            except Exception as e:

                                st.error(
                                    f"❌ Gagal menghapus "
                                    f"user: {e}"
                                )


        # =================================================
        # TAMBAH USER
        # =================================================

        with tab2:

            with st.container(
                border=True
            ):

                with st.form(
                    "add_user_form"
                ):

                    new_username = st.text_input(
                        "👤 Username Baru"
                    )

                    new_password = st.text_input(
                        "🔑 Password Baru",
                        type="password"
                    )

                    new_role = st.selectbox(
                        "🛡️ Hak Akses",
                        [
                            "user",
                            "admin"
                        ]
                    )

                    submit_new_user = (
                        st.form_submit_button(
                            "💾 Simpan Pengguna",
                            use_container_width=True
                        )
                    )


                    if submit_new_user:

                        if (
                            new_username
                            and new_password
                        ):

                            try:

                                # Cek username
                                existing = (
                                    supabase
                                    .table("users")
                                    .select(
                                        "username"
                                    )
                                    .eq(
                                        "username",
                                        new_username
                                    )
                                    .execute()
                                )

                                if existing.data:

                                    st.error(
                                        "⚠️ Username "
                                        "sudah terdaftar!"
                                    )

                                else:

                                    hashed_pw = (
                                        hash_password(
                                            new_password
                                        )
                                    )

                                    (
                                        supabase
                                        .table("users")
                                        .insert({
                                            "username":
                                                new_username,
                                            "password":
                                                hashed_pw,
                                            "role":
                                                new_role
                                        })
                                        .execute()
                                    )

                                    st.success(
                                        f"✅ Pengguna baru "
                                        f"**{new_username}** "
                                        f"berhasil "
                                        f"ditambahkan!"
                                    )

                                    st.rerun()

                            except Exception as e:

                                st.error(
                                    f"❌ Gagal menyimpan "
                                    f"user: {e}"
                                )

                        else:

                            st.warning(
                                "⚠️ Username dan "
                                "Password tidak boleh "
                                "kosong!"
                            )
