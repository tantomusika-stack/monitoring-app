"""
Sistem Monitoring BINPRES — Versi Lengkap
KONI Kabupaten Tangerang

Fitur:
- Login admin / petugas + ganti password
- Form input monitoring (5 section) + upload multi-foto
- Foto disimpan BLOB di database (aman di Streamlit Cloud)
- Dashboard admin + grafik statistik
- Filter laporan (cabor, status, tanggal, keyword)
- Detail laporan + preview foto
- Export Word (satuan & rekap, foto ikut) + Export Excel
- Laporan Saya (petugas): lihat status, detail, download Word
- Hapus laporan (admin)
- Manajemen user & cabor
- Download Word oleh user + arsip otomatis ke Google Drive
- Sinkronisasi ulang file Word ke Google Drive setelah tindak lanjut admin
"""

from __future__ import annotations

import streamlit as st
import datetime
import sqlite3
import hashlib
import uuid
import re
from io import BytesIO
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Google Drive API
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ============================================================
# KONFIGURASI
# ============================================================
st.set_page_config(
    page_title="Sistem Monitoring BINPRES",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_NAME = "monitoring.db"
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# ============================================================
# GOOGLE DRIVE
# ============================================================
# Folder tujuan arsip otomatis laporan Word.
GOOGLE_DRIVE_FOLDER_ID = "1AfjLphuOSf2ekqxCSv9OzU6x_B1qEvGf"
GOOGLE_DRIVE_SCOPE = ["https://www.googleapis.com/auth/drive"]
WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

DEFAULT_STATUS = "Belum Ditindaklanjuti"
STATUS_OPTIONS = [
    "Belum Ditindaklanjuti",
    "Sedang Ditindaklanjuti",
    "Selesai",
]

DEFAULT_CABOR = [
    "ANGGAR", "ANGKAT BERAT", "ANGKAT BESI", "AQUATIC/RENANG", "ARUNG JERAM",
    "ATLETIK", "BALAP SEPEDA", "BARONGSAI", "BERMOTOR", "BILLIARD", "BINARAGA",
    "BOLA BASKET", "BOLA TANGAN", "BOLA VOLI", "BOWLING", "BRIDGE", "BULUTANGKIS",
    "CATUR", "DAYUNG", "DRUMBAND", "E-SPORT", "FLOOR BALL", "FUTSAL", "GATEBALL",
    "GOLF", "GULAT", "GYMNASTIC/SENAM", "HOKI", "IBCA MMA", "JU JITSU", "JUDO",
    "KARATE", "KEMPO", "MENEMBAK", "MUAYTHAI", "PANAHAN", "PANJAT TEBING",
    "PENCAK SILAT", "PETANQUE", "PICKLEBALL", "RUGBY", "SAMBO", "SELAM",
    "SEPAK BOLA", "SEPAK TAKRAW", "SEPATU RODA", "SOFTBALL", "SQUASH",
    "TAEKWONDO", "TARUNG DERAJAT", "TENIS LAPANG", "TENIS MEJA", "TINJU",
    "WOODBALL", "WUSHU",
]

FIELD_LABELS: Dict[str, List[Tuple[str, str]]] = {
    "1. Performa Fisik & Kebugaran": [
        ("fisik_parameter", "Capaian parameter fisik (vs benchmark target)"),
        ("fisik_peaking", "Grafik performa puncak (peaking)"),
        ("fisik_recovery", "Tingkat pemulihan fisik (recovery)"),
        ("fisik_cedera", "Keluhan cedera lama / indikasi cedera baru"),
    ],
    "2. Kesiapan Taktis & Strategi": [
        ("taktis_lawan", "Pemetaan kekuatan calon lawan"),
        ("taktis_instruksi", "Kemampuan mengikuti instruksi teknis"),
        ("taktis_ujicoba", "Hasil try-out / sparing"),
    ],
    "3. Mental, Psikologis & Kesiapan Mental": [
        ("mental_cemas", "Tingkat kecemasan & pengendalian stres"),
        ("mental_fokus", "Fokus, motivasi, dan self-confidence"),
        ("mental_rutinitas", "Rutinitas mental khusus"),
        ("mental_psikolog", "Koordinasi dengan psikolog olahraga"),
    ],
    "4. Nutrisi, Berat Badan & Gaya Hidup": [
        ("nutrisi_bb", "Progres penyesuaian berat badan"),
        ("nutrisi_asupan", "Asupan nutrisi dan suplemen"),
        ("nutrisi_hidrasi", "Status hidrasi"),
        ("nutrisi_tidur", "Kualitas dan kecukupan tidur"),
    ],
    "5. Medis, Bebas Doping & Logistik": [
        ("medis_rekam", "Status rekam medis & tim medis"),
        ("medis_doping", "Keamanan obat / suplemen (bebas doping)"),
        ("medis_alat", "Kesiapan perlengkapan tanding"),
        ("medis_nonteknis", "Kendala non-teknis"),
    ],
}

CUSTOM_CSS = """
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.25);
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p { margin: 0.3rem 0 0 0; opacity: 0.9; font-size: 0.95rem; }

    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.8rem 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    div[data-testid="stMetric"] label { color: #64748b !important; font-size: 0.85rem !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #1e3a8a !important; font-weight: 700; }

    .section-badge {
        display: inline-block;
        background: #1e3a8a;
        color: white;
        padding: 0.35rem 0.9rem;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.95rem;
        margin-bottom: 0.8rem;
    }

    .badge-belum { background:#fef3c7; color:#92400e; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }
    .badge-sedang { background:#dbeafe; color:#1e40af; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }
    .badge-selesai { background:#d1fae5; color:#065f46; padding:3px 10px; border-radius:12px; font-size:0.8rem; font-weight:600; }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    }
    section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    section[data-testid="stSidebar"] .stButton > button {
        background: #334155; border: 1px solid #475569; color: white;
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: #475569; border-color: #64748b;
    }

    .stButton > button[kind="primary"],
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        margin-bottom: 0.5rem;
    }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
</style>
"""


# ============================================================
# DATABASE
# ============================================================
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, stored: str) -> bool:
    return hash_password(password) == stored or password == stored


def _add_column_if_missing(cursor: sqlite3.Cursor, table: str, column: str, definition: str) -> None:
    existing = [r["name"] for r in cursor.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in existing:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db() -> None:
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            nama_lengkap TEXT DEFAULT '',
            aktif INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS cabor (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT UNIQUE NOT NULL,
            aktif INTEGER DEFAULT 1
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS laporan_monitoring (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal DATE,
            cabor TEXT,
            lokasi TEXT,
            petugas TEXT,
            fisik_parameter TEXT,
            fisik_peaking TEXT,
            fisik_recovery TEXT,
            fisik_cedera TEXT,
            taktis_lawan TEXT,
            taktis_instruksi TEXT,
            taktis_ujicoba TEXT,
            mental_cemas TEXT,
            mental_fokus TEXT,
            mental_rutinitas TEXT,
            mental_psikolog TEXT,
            nutrisi_bb TEXT,
            nutrisi_asupan TEXT,
            nutrisi_hidrasi TEXT,
            nutrisi_tidur TEXT,
            medis_rekam TEXT,
            medis_doping TEXT,
            medis_alat TEXT,
            medis_nonteknis TEXT,
            status TEXT DEFAULT 'Belum Ditindaklanjuti',
            catatan_admin TEXT DEFAULT '',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS laporan_foto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            laporan_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            original_name TEXT,
            mime_type TEXT DEFAULT 'image/jpeg',
            data BLOB,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (laporan_id) REFERENCES laporan_monitoring(id) ON DELETE CASCADE
        )
    """)

    for col, defn in {
        "nama_lengkap": "TEXT DEFAULT ''",
        "aktif": "INTEGER DEFAULT 1",
        "created_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
    }.items():
        _add_column_if_missing(cur, "users", col, defn)

    for col, defn in {
        "status": "TEXT DEFAULT 'Belum Ditindaklanjuti'",
        "catatan_admin": "TEXT DEFAULT ''",
        "updated_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
        "created_at": "TEXT DEFAULT CURRENT_TIMESTAMP",
        "drive_file_id": "TEXT DEFAULT ''",
        "drive_url": "TEXT DEFAULT ''",
        "drive_status": "TEXT DEFAULT 'Belum diunggah'",
        "drive_uploaded_at": "TEXT DEFAULT ''",
    }.items():
        _add_column_if_missing(cur, "laporan_monitoring", col, defn)

    for col, defn in {
        "mime_type": "TEXT DEFAULT 'image/jpeg'",
        "data": "BLOB",
    }.items():
        _add_column_if_missing(cur, "laporan_foto", col, defn)

    count = cur.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if count == 0:
        cur.executemany(
            """INSERT INTO users (username, password, role, nama_lengkap, aktif)
               VALUES (?, ?, ?, ?, 1)""",
            [
                ("admin", hash_password("admin123"), "admin", "Administrator"),
                ("petugas", hash_password("petugas123"), "user", "Petugas Monitoring"),
            ],
        )
    else:
        for username, old_pw, role, nama in [
            ("admin", "admin123", "admin", "Administrator"),
            ("petugas", "petugas123", "user", "Petugas Monitoring"),
        ]:
            row = cur.execute(
                "SELECT id, password FROM users WHERE username = ?", (username,)
            ).fetchone()
            if row and row["password"] == old_pw:
                cur.execute(
                    "UPDATE users SET password=?, role=?, nama_lengkap=? WHERE id=?",
                    (hash_password(old_pw), role, nama, row["id"]),
                )

    if cur.execute("SELECT COUNT(*) AS n FROM cabor").fetchone()["n"] == 0:
        cur.executemany(
            "INSERT OR IGNORE INTO cabor (nama) VALUES (?)",
            [(n,) for n in DEFAULT_CABOR],
        )

    conn.commit()
    conn.close()


# ============================================================
# HELPER / QUERY
# ============================================================
def logout() -> None:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def get_cabor_list() -> List[str]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT nama FROM cabor WHERE aktif=1 ORDER BY nama"
        ).fetchall()
    return [r["nama"] for r in rows]


def get_current_user() -> Optional[sqlite3.Row]:
    username = st.session_state.get("username")
    if not username:
        return None
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM users WHERE username=?", (username,)
        ).fetchone()


def fetch_all_users() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(
            """SELECT id, username, nama_lengkap, role,
                      CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status,
                      created_at
               FROM users ORDER BY id DESC""",
            conn,
        )


def fetch_all_cabor() -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(
            """SELECT id, nama,
                      CASE WHEN aktif=1 THEN 'Aktif' ELSE 'Nonaktif' END AS status
               FROM cabor ORDER BY nama""",
            conn,
        )


def fetch_laporan_summary() -> Dict[str, int]:
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) AS n FROM laporan_monitoring").fetchone()["n"]
        bulan_ini = conn.execute(
            """SELECT COUNT(*) AS n FROM laporan_monitoring
               WHERE strftime('%Y-%m', tanggal)=strftime('%Y-%m','now')"""
        ).fetchone()["n"]
        cabor_termonitor = conn.execute(
            "SELECT COUNT(DISTINCT cabor) AS n FROM laporan_monitoring"
        ).fetchone()["n"]
        perlu_tindak = conn.execute(
            """SELECT COUNT(*) AS n FROM laporan_monitoring
               WHERE status=? OR status IS NULL""",
            (DEFAULT_STATUS,),
        ).fetchone()["n"]
        selesai = conn.execute(
            "SELECT COUNT(*) AS n FROM laporan_monitoring WHERE status='Selesai'"
        ).fetchone()["n"]
    return {
        "total": total,
        "bulan_ini": bulan_ini,
        "cabor_termonitor": cabor_termonitor,
        "perlu_tindak": perlu_tindak,
        "selesai": selesai,
    }


def fetch_chart_data():
    with get_conn() as conn:
        by_cabor = pd.read_sql_query(
            """SELECT cabor, COUNT(*) AS jumlah
               FROM laporan_monitoring
               GROUP BY cabor ORDER BY jumlah DESC LIMIT 15""",
            conn,
        )
        by_month = pd.read_sql_query(
            """SELECT strftime('%Y-%m', tanggal) AS bulan, COUNT(*) AS jumlah
               FROM laporan_monitoring
               WHERE tanggal IS NOT NULL
               GROUP BY bulan ORDER BY bulan""",
            conn,
        )
        by_status = pd.read_sql_query(
            """SELECT COALESCE(status, 'Belum Ditindaklanjuti') AS status, COUNT(*) AS jumlah
               FROM laporan_monitoring GROUP BY status""",
            conn,
        )
    return by_cabor, by_month, by_status


def fetch_laporan_list(
    petugas: Optional[str] = None,
    cabor: Optional[str] = None,
    status: Optional[str] = None,
    tgl_awal: Optional[datetime.date] = None,
    tgl_akhir: Optional[datetime.date] = None,
    keyword: Optional[str] = None,
) -> pd.DataFrame:
    clauses = []
    params: list = []

    if petugas:
        clauses.append("petugas = ?")
        params.append(petugas)
    if cabor and cabor != "Semua":
        clauses.append("cabor = ?")
        params.append(cabor)
    if status and status != "Semua":
        clauses.append("COALESCE(status, ?) = ?")
        params.extend([DEFAULT_STATUS, status])
    if tgl_awal:
        clauses.append("tanggal >= ?")
        params.append(str(tgl_awal))
    if tgl_akhir:
        clauses.append("tanggal <= ?")
        params.append(str(tgl_akhir))
    if keyword:
        clauses.append("(lokasi LIKE ? OR petugas LIKE ? OR cabor LIKE ?)")
        kw = f"%{keyword}%"
        params.extend([kw, kw, kw])

    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    sql = f"""
        SELECT id, tanggal, cabor, lokasi, petugas,
               COALESCE(status, '{DEFAULT_STATUS}') AS status,
               COALESCE(drive_status, 'Belum diunggah') AS drive_status,
               drive_url, updated_at
        FROM laporan_monitoring
        {where}
        ORDER BY tanggal DESC, id DESC
    """
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


def get_laporan_by_id(laporan_id: int) -> Optional[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM laporan_monitoring WHERE id=?", (laporan_id,)
        ).fetchone()


def get_fotos_by_laporan(laporan_id: int) -> List[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM laporan_foto WHERE laporan_id=? ORDER BY id",
            (laporan_id,),
        ).fetchall()


def _mime_from_ext(ext: str) -> str:
    return {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
    }.get(ext.lower(), "image/jpeg")


def get_foto_bytes(foto: sqlite3.Row) -> Optional[bytes]:
    """Prioritaskan BLOB di DB, fallback ke file disk."""
    try:
        blob = foto["data"]
        if blob:
            return bytes(blob)
    except (IndexError, KeyError):
        pass
    path = UPLOAD_DIR / foto["filename"]
    if path.exists():
        return path.read_bytes()
    return None


def save_uploaded_photos(laporan_id: int, uploaded_files: list) -> int:
    """Simpan foto ke disk + BLOB di database."""
    if not uploaded_files:
        return 0
    count = 0
    with get_conn() as conn:
        for f in uploaded_files:
            ext = Path(f.name).suffix.lower()
            if ext not in {".jpg", ".jpeg", ".png", ".webp"}:
                continue
            raw = f.getbuffer().tobytes()
            if not raw:
                continue
            unique_name = f"{laporan_id}_{uuid.uuid4().hex[:10]}{ext}"
            try:
                dest = UPLOAD_DIR / unique_name
                with open(dest, "wb") as out:
                    out.write(raw)
            except OSError:
                pass
            conn.execute(
                """INSERT INTO laporan_foto
                   (laporan_id, filename, original_name, mime_type, data)
                   VALUES (?, ?, ?, ?, ?)""",
                (laporan_id, unique_name, f.name, _mime_from_ext(ext), raw),
            )
            count += 1
        conn.commit()
    return count


def delete_laporan(laporan_id: int) -> None:
    fotos = get_fotos_by_laporan(laporan_id)
    for f in fotos:
        path = UPLOAD_DIR / f["filename"]
        if path.exists():
            try:
                path.unlink()
            except OSError:
                pass
    with get_conn() as conn:
        conn.execute("DELETE FROM laporan_foto WHERE laporan_id=?", (laporan_id,))
        conn.execute("DELETE FROM laporan_monitoring WHERE id=?", (laporan_id,))
        conn.commit()


def status_badge_html(status: str) -> str:
    s = status or DEFAULT_STATUS
    if s == "Selesai":
        cls = "badge-selesai"
    elif s == "Sedang Ditindaklanjuti":
        cls = "badge-sedang"
    else:
        cls = "badge-belum"
    return f'<span class="{cls}">{s}</span>'


# ============================================================
# GOOGLE DRIVE HELPERS
# ============================================================
def drive_is_configured() -> bool:
    """Cek apakah kredensial Google Drive tersedia di Streamlit Secrets."""
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def get_drive_service():
    """Buat koneksi Google Drive menggunakan service account dari st.secrets."""
    if not drive_is_configured():
        return None
    info = dict(st.secrets["gcp_service_account"])
    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=GOOGLE_DRIVE_SCOPE,
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def safe_drive_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r'[\\/:*?"<>|]+', "-", value)
    value = re.sub(r"\s+", " ", value)
    return value[:180] or "Laporan_Monitoring"


def get_drive_report_name(row: sqlite3.Row) -> str:
    return safe_drive_filename(
        f"Laporan_Monev_BINPRES_ID-{row['id']}_{row['cabor']}_{row['tanggal']}.docx"
    )


def upload_or_update_drive_report(laporan_id: int) -> Tuple[bool, str, str]:
    """Generate Word terbaru lalu upload/update ke folder Google Drive."""
    row = get_laporan_by_id(int(laporan_id))
    if not row:
        return False, "Laporan tidak ditemukan.", ""

    if not drive_is_configured():
        return False, "Google Drive belum dikonfigurasi di Streamlit Secrets.", ""

    service = get_drive_service()
    if service is None:
        return False, "Koneksi Google Drive tidak tersedia.", ""

    try:
        word_bytes = generate_word_report([dict(row)], is_all=False, include_photos=True)
        filename = get_drive_report_name(row)
        media = MediaIoBaseUpload(
            BytesIO(word_bytes),
            mimetype=WORD_MIME,
            resumable=True,
        )

        existing_id = (row["drive_file_id"] or "").strip()
        if existing_id:
            try:
                result = service.files().update(
                    fileId=existing_id,
                    body={"name": filename},
                    media_body=media,
                    fields="id,webViewLink,name",
                ).execute()
                file_id = result["id"]
            except Exception:
                # Jika file lama sudah dihapus/dipindahkan dan ID tidak valid,
                # buat arsip baru agar laporan tetap tersimpan.
                existing_id = ""

        if not existing_id:
            metadata = {
                "name": filename,
                "parents": [GOOGLE_DRIVE_FOLDER_ID],
                "mimeType": WORD_MIME,
                "description": (
                    f"Laporan Monitoring BINPRES KONI Kabupaten Tangerang | "
                    f"ID {row['id']} | Cabor {row['cabor']} | Tanggal {row['tanggal']}"
                ),
            }
            result = service.files().create(
                body=metadata,
                media_body=media,
                fields="id,webViewLink,name",
            ).execute()
            file_id = result["id"]

        url = result.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_conn() as conn:
            conn.execute(
                """UPDATE laporan_monitoring
                   SET drive_file_id=?, drive_url=?, drive_status=?, drive_uploaded_at=?
                   WHERE id=?""",
                (file_id, url, "Tersimpan di Google Drive", now, int(laporan_id)),
            )
            conn.commit()
        return True, "Laporan berhasil diarsipkan ke Google Drive.", url

    except Exception as exc:
        message = str(exc)
        with get_conn() as conn:
            conn.execute(
                """UPDATE laporan_monitoring
                   SET drive_status=?
                   WHERE id=?""",
                (f"Gagal: {message[:180]}", int(laporan_id)),
            )
            conn.commit()
        return False, f"Upload Google Drive gagal: {message}", ""


def retry_drive_upload(laporan_id: int) -> Tuple[bool, str, str]:
    return upload_or_update_drive_report(int(laporan_id))


# ============================================================
# WORD & EXCEL EXPORT
# ============================================================
def _set_run_font(run, size_pt: float = 10, bold: bool = False) -> None:
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.name = "Calibri"


def _add_compact_para(doc, text: str, bold: bool = False, size: float = 10, space_after: float = 2) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    _set_run_font(run, size_pt=size, bold=bold)


def generate_word_report(
    data_list: List[Dict[str, Any]],
    is_all: bool = False,
    include_photos: bool = True,
) -> bytes:
    """
    Layout hemat kertas:
    - Margin kecil, font 10pt, spasi rapat
    - Lembar isi laporan dulu
    - Lembar foto terpisah (halaman baru) di akhir tiap laporan
    - 1 file Word
    """
    doc = Document()
    section = doc.sections[0]
    # Margin rapat (hemat kertas)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # Style default body
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(2)
    style.paragraph_format.line_spacing = 1.0

    judul = "REKAPITULASI LAPORAN MONEV BINPRES" if is_all else "LAPORAN MONEV BINPRES"
    title = doc.add_heading(judul, level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(14)
        run.font.name = "Calibri"

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.paragraph_format.space_after = Pt(6)
    r = meta.add_run(
        f"KONI Kabupaten Tangerang  |  Dicetak: {datetime.date.today().strftime('%d/%m/%Y')}"
    )
    _set_run_font(r, size_pt=9)

    def add_info_line(label: str, value: Any) -> None:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        r1 = p.add_run(f"{label}: ")
        _set_run_font(r1, size_pt=10, bold=True)
        r2 = p.add_run(str(value) if value else "-")
        _set_run_font(r2, size_pt=10)

    def add_section_compact(title_text: str, items: List[Tuple[str, Any]]) -> None:
        # Judul section
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(2)
        hr = h.add_run(title_text)
        _set_run_font(hr, size_pt=11, bold=True)

        for question, answer in items:
            # Pertanyaan + jawaban dalam 1 blok rapat
            pq = doc.add_paragraph()
            pq.paragraph_format.space_before = Pt(2)
            pq.paragraph_format.space_after = Pt(0)
            pq.paragraph_format.line_spacing = 1.0
            rq = pq.add_run(f"• {question}")
            _set_run_font(rq, size_pt=9, bold=True)

            pa = doc.add_paragraph()
            pa.paragraph_format.space_before = Pt(0)
            pa.paragraph_format.space_after = Pt(2)
            pa.paragraph_format.line_spacing = 1.0
            pa.paragraph_format.left_indent = Inches(0.15)
            ra = pa.add_run(str(answer).strip() if answer and str(answer).strip() else "—")
            _set_run_font(ra, size_pt=9)

    # Kumpulkan foto per laporan untuk halaman terpisah
    photos_queue: List[Tuple[Dict[str, Any], list]] = []

    for idx, data in enumerate(data_list):
        if is_all and idx > 0:
            doc.add_page_break()

        if is_all:
            h = doc.add_paragraph()
            h.paragraph_format.space_before = Pt(4)
            h.paragraph_format.space_after = Pt(4)
            hr = h.add_run(
                f"Laporan {idx + 1}: {data.get('cabor', '-')} — {data.get('tanggal', '-')}"
            )
            _set_run_font(hr, size_pt=12, bold=True)

        # --- Lembar isi (rapat) ---
        add_info_line("Cabang Olahraga", data.get("cabor"))
        add_info_line("Tanggal", data.get("tanggal"))
        add_info_line("Lokasi", data.get("lokasi"))
        add_info_line("Petugas Monev", data.get("petugas"))
        add_info_line("Status", data.get("status") or DEFAULT_STATUS)

        # Garis pemisah tipis
        sep = doc.add_paragraph()
        sep.paragraph_format.space_before = Pt(2)
        sep.paragraph_format.space_after = Pt(2)
        sr = sep.add_run("─" * 55)
        _set_run_font(sr, size_pt=8)

        for section_title, fields in FIELD_LABELS.items():
            items = [(label, data.get(key)) for key, label in fields]
            add_section_compact(section_title, items)

        if data.get("catatan_admin"):
            add_section_compact("Catatan Admin / Tindak Lanjut", [
                ("Catatan", data["catatan_admin"]),
            ])

        # Simpan foto untuk halaman terpisah
        if include_photos and data.get("id"):
            fotos = get_fotos_by_laporan(int(data["id"]))
            if fotos:
                photos_queue.append((data, fotos))

    # --- Lembar foto terpisah (di akhir, tetap 1 file) ---
    if include_photos and photos_queue:
        for data, fotos in photos_queue:
            doc.add_page_break()

            h = doc.add_paragraph()
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_after = Pt(4)
            hr = h.add_run("DOKUMENTASI FOTO")
            _set_run_font(hr, size_pt=12, bold=True)

            sub = doc.add_paragraph()
            sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sub.paragraph_format.space_after = Pt(8)
            sr = sub.add_run(
                f"{data.get('cabor', '-')}  |  {data.get('tanggal', '-')}  |  "
                f"{data.get('lokasi', '-')}  |  Petugas: {data.get('petugas', '-')}"
            )
            _set_run_font(sr, size_pt=9)

            for foto in fotos:
                img_bytes = get_foto_bytes(foto)
                if img_bytes:
                    try:
                        # Lebar sedang agar hemat ruang, bisa 2 foto per halaman
                        doc.add_picture(BytesIO(img_bytes), width=Inches(5.2))
                        cap = doc.add_paragraph()
                        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cap.paragraph_format.space_before = Pt(2)
                        cap.paragraph_format.space_after = Pt(8)
                        cr = cap.add_run(foto["original_name"] or foto["filename"])
                        _set_run_font(cr, size_pt=8)
                    except Exception:
                        _add_compact_para(
                            doc, f"[Gagal memuat: {foto['original_name']}]", size=9
                        )
                else:
                    _add_compact_para(
                        doc, f"[Foto tidak tersedia: {foto['original_name']}]", size=9
                    )

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generate_excel_report(df: pd.DataFrame) -> bytes:
    if df.empty:
        out = BytesIO()
        pd.DataFrame({"info": ["Tidak ada data"]}).to_excel(out, index=False)
        return out.getvalue()

    ids = df["id"].tolist()
    placeholders = ",".join("?" * len(ids))
    with get_conn() as conn:
        detail = pd.read_sql_query(
            f"""SELECT * FROM laporan_monitoring
                WHERE id IN ({placeholders})
                ORDER BY tanggal DESC, id DESC""",
            conn,
            params=ids,
        )

    rename_map = {
        "id": "ID",
        "tanggal": "Tanggal",
        "cabor": "Cabang Olahraga",
        "lokasi": "Lokasi",
        "petugas": "Petugas",
        "status": "Status",
        "catatan_admin": "Catatan Admin",
        "fisik_parameter": "Fisik - Parameter",
        "fisik_peaking": "Fisik - Peaking",
        "fisik_recovery": "Fisik - Recovery",
        "fisik_cedera": "Fisik - Cedera",
        "taktis_lawan": "Taktis - Lawan",
        "taktis_instruksi": "Taktis - Instruksi",
        "taktis_ujicoba": "Taktis - Ujicoba",
        "mental_cemas": "Mental - Cemas",
        "mental_fokus": "Mental - Fokus",
        "mental_rutinitas": "Mental - Rutinitas",
        "mental_psikolog": "Mental - Psikolog",
        "nutrisi_bb": "Nutrisi - BB",
        "nutrisi_asupan": "Nutrisi - Asupan",
        "nutrisi_hidrasi": "Nutrisi - Hidrasi",
        "nutrisi_tidur": "Nutrisi - Tidur",
        "medis_rekam": "Medis - Rekam",
        "medis_doping": "Medis - Doping",
        "medis_alat": "Medis - Alat",
        "medis_nonteknis": "Medis - Nonteknis",
        "updated_at": "Diperbarui",
        "created_at": "Dibuat",
        "drive_status": "Status Google Drive",
        "drive_url": "Link Google Drive",
        "drive_uploaded_at": "Waktu Upload Drive",
    }
    detail = detail.rename(
        columns={k: v for k, v in rename_map.items() if k in detail.columns}
    )

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Laporan Monitoring", index=False)
    return out.getvalue()


# ============================================================
# UI COMPONENTS
# ============================================================
def render_sidebar() -> None:
    user = get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES")
        st.caption("Monitoring & Evaluasi Cabor")
        st.markdown("---")
        nama = (
            user["nama_lengkap"]
            if user and user["nama_lengkap"]
            else st.session_state.get("username", "")
        )
        st.write(f"👤 **{nama}**")
        st.caption(f"Role: {st.session_state.get('role', '').upper()}")
        st.markdown("---")

        with st.expander("🔑 Ganti Password"):
            with st.form("form_ganti_password"):
                pw_lama = st.text_input("Password lama", type="password")
                pw_baru = st.text_input("Password baru", type="password")
                pw_konfirm = st.text_input("Konfirmasi password baru", type="password")
                if st.form_submit_button("Simpan Password", use_container_width=True):
                    if not pw_lama or not pw_baru:
                        st.error("Semua field wajib diisi.")
                    elif len(pw_baru) < 6:
                        st.error("Password baru minimal 6 karakter.")
                    elif pw_baru != pw_konfirm:
                        st.error("Konfirmasi password tidak cocok.")
                    else:
                        u = get_current_user()
                        if u and verify_password(pw_lama, u["password"]):
                            with get_conn() as conn:
                                conn.execute(
                                    "UPDATE users SET password=? WHERE id=?",
                                    (hash_password(pw_baru), u["id"]),
                                )
                                conn.commit()
                            st.success("Password berhasil diubah.")
                        else:
                            st.error("Password lama salah.")

        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            logout()


def halaman_login() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div style="text-align:center;padding:40px 0 10px 0">
            <div style="font-size:56px">🏆</div>
            <h1 style="margin-bottom:4px;color:#1e3a8a">Sistem Monitoring BINPRES</h1>
            <p style="color:#64748b;font-size:1.05rem">KONI Kabupaten Tangerang</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, col_c, _ = st.columns([1, 1.4, 1])
    with col_c:
        with st.form("login_form"):
            st.markdown("#### Masuk ke Sistem")
            username = st.text_input("👤 Username", placeholder="Masukkan username")
            password = st.text_input(
                "🔒 Password", type="password", placeholder="Masukkan password"
            )
            submit = st.form_submit_button(
                "Masuk ke Sistem", use_container_width=True, type="primary"
            )

            if submit:
                with get_conn() as conn:
                    user = conn.execute(
                        "SELECT * FROM users WHERE username=? AND aktif=1",
                        (username.strip(),),
                    ).fetchone()

                if user and verify_password(password, user["password"]):
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = user["username"]
                    st.session_state["role"] = user["role"]
                    st.session_state["nama_lengkap"] = (
                        user["nama_lengkap"] or user["username"]
                    )
                    st.rerun()
                else:
                    st.error("Username atau Password salah / akun tidak aktif.")

        st.caption("Data tersimpan pada database SQLite lokal • Versi lengkap")


def render_detail_laporan(row: sqlite3.Row) -> None:
    st.markdown(
        f"**Cabor:** {row['cabor']} &nbsp;|&nbsp; "
        f"**Tanggal:** {row['tanggal']} &nbsp;|&nbsp; "
        f"**Lokasi:** {row['lokasi']} &nbsp;|&nbsp; "
        f"**Petugas:** {row['petugas']}"
    )
    st.markdown(
        f"**Status:** {status_badge_html(row['status'] or DEFAULT_STATUS)}",
        unsafe_allow_html=True,
    )
    if row["catatan_admin"]:
        st.info(f"**Catatan Admin:** {row['catatan_admin']}")

    for section_title, fields in FIELD_LABELS.items():
        with st.expander(section_title, expanded=False):
            for key, label in fields:
                val = row[key] if key in row.keys() else None
                st.markdown(f"**{label}**")
                st.write(val if val else "—")

    fotos = get_fotos_by_laporan(int(row["id"]))
    if fotos:
        st.markdown("#### 🖼️ Dokumentasi Foto")
        cols = st.columns(min(4, len(fotos)))
        for i, foto in enumerate(fotos):
            img_bytes = get_foto_bytes(foto)
            if img_bytes:
                with cols[i % len(cols)]:
                    st.image(
                        img_bytes,
                        caption=foto["original_name"],
                        use_container_width=True,
                    )


# ------------------------------------------------------------
# DASHBOARD ADMIN
# ------------------------------------------------------------
def dashboard_admin() -> None:
    stats = fetch_laporan_summary()

    st.markdown(
        """
        <div class="main-header">
            <h1>📊 Dashboard Monitoring BINPRES</h1>
            <p>Ringkasan aktivitas monitoring & evaluasi cabang olahraga</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Laporan", stats["total"])
    c2.metric("Bulan Ini", stats["bulan_ini"])
    c3.metric("Cabor Termonitor", stats["cabor_termonitor"])
    c4.metric("Perlu Tindak Lanjut", stats["perlu_tindak"])
    c5.metric("Selesai", stats["selesai"])

    st.markdown("---")

    by_cabor, by_month, by_status = fetch_chart_data()

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("##### 🏅 Top Cabor (jumlah laporan)")
        if not by_cabor.empty:
            st.bar_chart(by_cabor.set_index("cabor")["jumlah"], height=280)
        else:
            st.caption("Belum ada data.")

    with col_b:
        st.markdown("##### 📅 Tren Laporan per Bulan")
        if not by_month.empty:
            st.line_chart(by_month.set_index("bulan")["jumlah"], height=280)
        else:
            st.caption("Belum ada data.")

    st.markdown("##### 📌 Distribusi Status")
    if not by_status.empty:
        st.bar_chart(by_status.set_index("status")["jumlah"], height=200)
    else:
        st.caption("Belum ada data.")

    st.markdown("---")


# ------------------------------------------------------------
# MANAJEMEN USER
# ------------------------------------------------------------
def kelola_user() -> None:
    st.subheader("👥 Manajemen User")
    st.dataframe(fetch_all_users(), use_container_width=True, hide_index=True)

    st.markdown("### ➕ Tambah User")
    with st.form("tambah_user"):
        col1, col2 = st.columns(2)
        with col1:
            username = st.text_input("Username baru")
            nama = st.text_input("Nama lengkap")
        with col2:
            password = st.text_input("Password", type="password")
            role = st.selectbox("Hak akses", ["user", "admin"])

        if st.form_submit_button("Simpan User", use_container_width=True, type="primary"):
            if not username.strip() or not password:
                st.error("Username dan password wajib diisi.")
            elif len(password) < 6:
                st.error("Password minimal 6 karakter.")
            else:
                try:
                    with get_conn() as conn:
                        conn.execute(
                            """INSERT INTO users
                               (username, password, role, nama_lengkap, aktif)
                               VALUES (?, ?, ?, ?, 1)""",
                            (
                                username.strip(),
                                hash_password(password),
                                role,
                                nama.strip(),
                            ),
                        )
                        conn.commit()
                    st.success("User berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Username sudah digunakan.")

    st.markdown("### ✏️ Edit / Nonaktifkan User")
    with get_conn() as conn:
        users = [
            dict(r)
            for r in conn.execute(
                "SELECT id, username, nama_lengkap, role, aktif FROM users ORDER BY username"
            ).fetchall()
        ]

    if not users:
        return

    pilihan = st.selectbox(
        "Pilih user",
        users,
        format_func=lambda x: f"{x['username']} — {x['nama_lengkap'] or '-'}",
    )

    with st.form("edit_user"):
        nama_baru = st.text_input("Nama lengkap", value=pilihan["nama_lengkap"] or "")
        role_baru = st.selectbox(
            "Role",
            ["user", "admin"],
            index=0 if pilihan["role"] == "user" else 1,
        )
        aktif_baru = st.checkbox("Akun aktif", value=bool(pilihan["aktif"]))
        password_baru = st.text_input(
            "Password baru (kosongkan jika tidak diubah)", type="password"
        )

        if st.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
            with get_conn() as conn:
                if password_baru:
                    if len(password_baru) < 6:
                        st.error("Password minimal 6 karakter.")
                        return
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap=?, role=?, aktif=?, password=?
                           WHERE id=?""",
                        (
                            nama_baru.strip(),
                            role_baru,
                            int(aktif_baru),
                            hash_password(password_baru),
                            pilihan["id"],
                        ),
                    )
                else:
                    conn.execute(
                        """UPDATE users
                           SET nama_lengkap=?, role=?, aktif=?
                           WHERE id=?""",
                        (
                            nama_baru.strip(),
                            role_baru,
                            int(aktif_baru),
                            pilihan["id"],
                        ),
                    )
                conn.commit()
            st.success("Data user berhasil diperbarui.")
            st.rerun()


# ------------------------------------------------------------
# MANAJEMEN CABOR
# ------------------------------------------------------------
def kelola_cabor() -> None:
    st.subheader("🏅 Kelola Cabang Olahraga")
    st.dataframe(fetch_all_cabor(), use_container_width=True, hide_index=True)

    with st.form("tambah_cabor"):
        nama_cabor = st.text_input("Nama Cabang Olahraga Baru")
        if st.form_submit_button("➕ Tambah Cabor", use_container_width=True):
            if not nama_cabor.strip():
                st.error("Nama cabor wajib diisi.")
            else:
                try:
                    with get_conn() as conn:
                        conn.execute(
                            "INSERT INTO cabor (nama) VALUES (?)",
                            (nama_cabor.strip().upper(),),
                        )
                        conn.commit()
                    st.success("Cabor berhasil ditambahkan.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Cabor tersebut sudah ada.")

    st.markdown("### ✏️ Edit / Nonaktifkan Cabor")
    with get_conn() as conn:
        rows = [
            dict(r)
            for r in conn.execute("SELECT * FROM cabor ORDER BY nama").fetchall()
        ]

    if not rows:
        return

    pilihan = st.selectbox("Pilih cabor", rows, format_func=lambda x: x["nama"])
    with st.form("edit_cabor"):
        nama_baru = st.text_input("Nama cabor", value=pilihan["nama"])
        aktif = st.checkbox("Cabor aktif", value=bool(pilihan["aktif"]))
        if st.form_submit_button("💾 Simpan", use_container_width=True):
            try:
                with get_conn() as conn:
                    conn.execute(
                        "UPDATE cabor SET nama=?, aktif=? WHERE id=?",
                        (nama_baru.strip().upper(), int(aktif), pilihan["id"]),
                    )
                    conn.commit()
                st.success("Cabor berhasil diperbarui.")
                st.rerun()
            except sqlite3.IntegrityError:
                st.error("Nama cabor sudah digunakan.")


# ------------------------------------------------------------
# LAPORAN ADMIN
# ------------------------------------------------------------
def halaman_laporan_admin() -> None:
    st.subheader("📄 Data Laporan Monitoring")

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        cabor_opts = ["Semua"] + get_cabor_list()
        cabor_filter = st.selectbox("Filter Cabor", cabor_opts)
    with f2:
        status_filter = st.selectbox("Filter Status", ["Semua"] + STATUS_OPTIONS)
    with f3:
        tgl_awal = st.date_input("Dari tanggal", value=None)
    with f4:
        tgl_akhir = st.date_input("Sampai tanggal", value=None)

    keyword = st.text_input(
        "🔎 Cari lokasi / petugas / cabor", placeholder="Ketik kata kunci..."
    )

    df = fetch_laporan_list(
        cabor=cabor_filter,
        status=status_filter,
        tgl_awal=tgl_awal if tgl_awal else None,
        tgl_akhir=tgl_akhir if tgl_akhir else None,
        keyword=keyword.strip() or None,
    )

    st.caption(f"Menampilkan **{len(df)}** laporan.")
    if df.empty:
        st.info("Tidak ada laporan sesuai filter.")
        return

    st.dataframe(df, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------
    # DOWNLOAD WORD PER LAPORAN
    # ------------------------------------------------------------
    st.markdown("### ⬇️ Download Word per Laporan")
    st.caption("Setiap laporan yang sudah tersimpan memiliki tombol Download Word masing-masing.")

    for _, item in df.iterrows():
        laporan_id = int(item["id"])
        status_item = item["status"] or DEFAULT_STATUS
        c_info, c_download = st.columns([5, 1.3])

        with c_info:
            st.markdown(
                f"**ID {laporan_id} — {item['cabor']}**  \n"
                f"📅 {item['tanggal']} &nbsp;|&nbsp; 📍 {item['lokasi']} &nbsp;|&nbsp; "
                f"👤 {item['petugas']} &nbsp;|&nbsp; **{status_item}**"
            )

        with c_download:
            laporan_row = get_laporan_by_id(laporan_id)
            if laporan_row:
                word_per_laporan = generate_word_report(
                    [dict(laporan_row)],
                    is_all=False,
                    include_photos=True,
                )
                st.download_button(
                    "⬇️ Download Word",
                    data=word_per_laporan,
                    file_name=f"Laporan_Monev_BINPRES_ID-{laporan_id}_{item['cabor']}_{item['tanggal']}.docx",
                    mime=WORD_MIME,
                    use_container_width=True,
                    key=f"download_word_admin_{laporan_id}",
                )

        st.markdown("---")

    st.markdown("### 🔍 Detail, Export & Tindak Lanjut")

    pilihan_id = st.selectbox(
        "Pilih laporan",
        df["id"].tolist(),
        format_func=lambda x: (
            f"ID {x} — {df.loc[df['id']==x, 'cabor'].values[0]} "
            f"({df.loc[df['id']==x, 'tanggal'].values[0]}) — "
            f"{df.loc[df['id']==x, 'status'].values[0]}"
        ),
        key="admin_pilih_laporan",
    )

    row = get_laporan_by_id(int(pilihan_id))
    if not row:
        st.error("Laporan tidak ditemukan.")
        return

    with st.expander("📋 Lihat Detail Lengkap", expanded=False):
        render_detail_laporan(row)

    if row["drive_url"]:
        st.success(f"☁️ Arsip Drive: {row['drive_status']} — [Buka file di Google Drive]({row['drive_url']})")
    else:
        st.warning(f"☁️ Arsip Drive: {row['drive_status'] or 'Belum diunggah'}")

    col_w, col_x, col_g, col_d = st.columns(4)
    with col_w:
        word_file = generate_word_report([dict(row)], is_all=False)
        st.download_button(
            "⬇️ Word Satuan",
            data=word_file,
            file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
            mime=WORD_MIME,
            use_container_width=True,
        )
    with col_x:
        excel_file = generate_excel_report(df)
        st.download_button(
            "⬇️ Excel (hasil filter)",
            data=excel_file,
            file_name=f"Rekap_Monev_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_g:
        if row["drive_url"]:
            st.link_button("☁️ Buka di Google Drive", row["drive_url"], use_container_width=True)
        else:
            if st.button("☁️ Upload ke Drive", use_container_width=True):
                ok, msg, url = retry_drive_upload(int(pilihan_id))
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    with col_d:
        if st.button("🗑️ Hapus Laporan", use_container_width=True):
            st.session_state["confirm_delete_id"] = int(pilihan_id)

    if st.session_state.get("confirm_delete_id") == int(pilihan_id):
        st.warning(
            f"Yakin hapus laporan ID {pilihan_id}? Tindakan ini tidak dapat dibatalkan."
        )
        c_yes, c_no = st.columns(2)
        with c_yes:
            if st.button("✅ Ya, Hapus", type="primary", use_container_width=True):
                delete_laporan(int(pilihan_id))
                st.session_state.pop("confirm_delete_id", None)
                st.success("Laporan berhasil dihapus.")
                st.rerun()
        with c_no:
            if st.button("❌ Batal", use_container_width=True):
                st.session_state.pop("confirm_delete_id", None)
                st.rerun()

    st.markdown("#### 📚 Export Rekap Word (semua hasil filter)")
    if st.button("Generate Word Rekap", use_container_width=True):
        ids = tuple(int(x) for x in df["id"].tolist())
        placeholders = ",".join("?" * len(ids))
        with get_conn() as conn:
            rows = conn.execute(
                f"""SELECT * FROM laporan_monitoring
                    WHERE id IN ({placeholders})
                    ORDER BY tanggal ASC, id ASC""",
                ids,
            ).fetchall()
        word_rekap = generate_word_report([dict(r) for r in rows], is_all=True)
        st.download_button(
            "⬇️ Download Rekap Word",
            data=word_rekap,
            file_name=f"Rekap_Monev_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("### ✏️ Tindak Lanjut Laporan")
    with st.form("update_status_laporan"):
        current = row["status"] or DEFAULT_STATUS
        status = st.selectbox(
            "Status tindak lanjut",
            STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(current) if current in STATUS_OPTIONS else 0,
        )
        catatan = st.text_area(
            "Catatan Admin / Rekomendasi",
            value=row["catatan_admin"] or "",
            height=100,
        )
        if st.form_submit_button("💾 Simpan Tindak Lanjut", use_container_width=True):
            with get_conn() as conn:
                conn.execute(
                    """UPDATE laporan_monitoring
                       SET status=?, catatan_admin=?, updated_at=CURRENT_TIMESTAMP
                       WHERE id=?""",
                    (status, catatan, int(pilihan_id)),
                )
                conn.commit()
            ok_drive, msg_drive, _ = upload_or_update_drive_report(int(pilihan_id))
            if ok_drive:
                st.success("Tindak lanjut berhasil diperbarui dan Word di Google Drive sudah diperbarui.")
            else:
                st.warning(f"Tindak lanjut tersimpan, tetapi sinkronisasi Drive belum berhasil: {msg_drive}")
            st.rerun()


# ------------------------------------------------------------
# HALAMAN ADMIN
# ------------------------------------------------------------
def halaman_admin() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_sidebar()
    dashboard_admin()

    tab1, tab2, tab3 = st.tabs([
        "📄 Laporan Monitoring",
        "👥 Manajemen User",
        "🏅 Kelola Cabor",
    ])
    with tab1:
        halaman_laporan_admin()
    with tab2:
        kelola_user()
    with tab3:
        kelola_cabor()


# ------------------------------------------------------------
# HALAMAN PETUGAS
# ------------------------------------------------------------
def form_input_monitoring() -> None:
    st.markdown(
        """
        <div class="main-header">
            <h1>📝 Form Input Monitoring</h1>
            <p>Isi laporan monitoring cabang olahraga secara lengkap</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        f"Petugas: **{st.session_state.get('nama_lengkap', st.session_state['username'])}**"
    )

    cabor_list = get_cabor_list()

    # ----------------------------------------------------------
    # UPLOAD FOTO DI LUAR FORM
    # (file_uploader di dalam form sering tidak muncul di Cloud)
    # ----------------------------------------------------------
    st.markdown(
        '<span class="section-badge">📷 Dokumentasi Foto</span>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Upload foto kegiatan **sebelum** mengisi form di bawah. Bisa lebih dari satu foto."
    )
    uploaded_files = st.file_uploader(
        "Pilih foto (JPG / PNG / WEBP)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="uploader_monitoring",
        help="Foto akan ikut tersimpan di laporan dan muncul di file Word saat diunduh.",
    )
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} foto siap diunggah bersama laporan.")
        preview_cols = st.columns(min(4, len(uploaded_files)))
        for i, f in enumerate(uploaded_files):
            with preview_cols[i % len(preview_cols)]:
                st.image(f, caption=f.name, use_container_width=True)
    else:
        st.info(
            "Belum ada foto dipilih. Foto opsional, tapi disarankan untuk dokumentasi."
        )

    st.markdown("---")

    with st.form("form_monitoring", clear_on_submit=True):
        st.markdown(
            '<span class="section-badge">📌 Informasi Dasar</span>',
            unsafe_allow_html=True,
        )
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Monitoring", datetime.date.today())
            cabor = st.selectbox("Cabang Olahraga", ["Pilih Cabor..."] + cabor_list)
        with col2:
            lokasi = st.text_input(
                "Lokasi Latihan / Try-out", placeholder="Contoh: GOR Cikokol"
            )

        st.markdown("---")

        with st.expander("💪 1. Performa Fisik & Kebugaran", expanded=True):
            fisik_1 = st.text_area(
                "Capaian parameter fisik (vs benchmark target):", height=70
            )
            fisik_2 = st.text_area(
                "Apakah atlet mencapai grafik performa puncak (peaking)?", height=70
            )
            fisik_3 = st.text_area("Tingkat pemulihan fisik (recovery):", height=70)
            fisik_4 = st.text_area(
                "Keluhan cedera lama / indikasi cedera baru:", height=70
            )

        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1 = st.text_area("Pemetaan kekuatan calon lawan:", height=70)
            taktis_2 = st.text_area(
                "Kemampuan mengikuti instruksi teknis di bawah tekanan:", height=70
            )
            taktis_3 = st.text_area("Hasil try-out / sparing:", height=70)

        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_1 = st.text_area(
                "Tingkat kecemasan & kemampuan mengendalikan stres:", height=70
            )
            mental_2 = st.text_area(
                "Fokus, motivasi, dan self-confidence:", height=70
            )
            mental_3 = st.text_area(
                "Rutinitas mental khusus saat bertanding:", height=70
            )
            mental_4 = st.text_area(
                "Koordinasi dengan tim psikolog olahraga:", height=70
            )

        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_1 = st.text_area("Progres penyesuaian berat badan:", height=70)
            nutrisi_2 = st.text_area(
                "Pemantauan asupan nutrisi dan suplemen:", height=70
            )
            nutrisi_3 = st.text_area("Status hidrasi harian atlet:", height=70)
            nutrisi_4 = st.text_area(
                "Kualitas dan kecukupan waktu tidur:", height=70
            )

        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1 = st.text_area(
                "Status rekam medis terkini & kesiapan fisioterapis:", height=70
            )
            medis_2 = st.text_area(
                "Keamanan obat, suplemen (Bebas Doping):", height=70
            )
            medis_3 = st.text_area(
                "Kesiapan perlengkapan khusus bertanding:", height=70
            )
            medis_4 = st.text_area(
                "Kendala non-teknis (akomodasi, transportasi):", height=70
            )

        st.markdown("---")
        n_ready = len(uploaded_files) if uploaded_files else 0
        if n_ready:
            st.caption(f"📷 {n_ready} foto akan disimpan bersama laporan ini.")
        else:
            st.caption(
                "📷 Tidak ada foto yang dipilih (laporan tetap bisa disimpan tanpa foto)."
            )

        submitted = st.form_submit_button(
            "💾 Simpan Laporan Monitoring",
            use_container_width=True,
            type="primary",
        )

        if submitted:
            if cabor == "Pilih Cabor...":
                st.error("⚠️ Harap pilih Cabang Olahraga.")
            elif not lokasi.strip():
                st.error("⚠️ Lokasi wajib diisi.")
            else:
                with get_conn() as conn:
                    cur = conn.execute(
                        """INSERT INTO laporan_monitoring (
                            tanggal, cabor, lokasi, petugas,
                            fisik_parameter, fisik_peaking, fisik_recovery, fisik_cedera,
                            taktis_lawan, taktis_instruksi, taktis_ujicoba,
                            mental_cemas, mental_fokus, mental_rutinitas, mental_psikolog,
                            nutrisi_bb, nutrisi_asupan, nutrisi_hidrasi, nutrisi_tidur,
                            medis_rekam, medis_doping, medis_alat, medis_nonteknis,
                            status, catatan_admin
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            tanggal,
                            cabor,
                            lokasi.strip(),
                            st.session_state["username"],
                            fisik_1,
                            fisik_2,
                            fisik_3,
                            fisik_4,
                            taktis_1,
                            taktis_2,
                            taktis_3,
                            mental_1,
                            mental_2,
                            mental_3,
                            mental_4,
                            nutrisi_1,
                            nutrisi_2,
                            nutrisi_3,
                            nutrisi_4,
                            medis_1,
                            medis_2,
                            medis_3,
                            medis_4,
                            DEFAULT_STATUS,
                            "",
                        ),
                    )
                    laporan_id = cur.lastrowid
                    conn.commit()

                files_to_save = (
                    st.session_state.get("uploader_monitoring")
                    or uploaded_files
                    or []
                )
                n_foto = save_uploaded_photos(laporan_id, files_to_save)

                # Otomatis buat Word dan arsipkan ke Google Drive.
                ok_drive, msg_drive, drive_url = upload_or_update_drive_report(int(laporan_id))

                msg = f"✅ Laporan berhasil disimpan (ID: {laporan_id})."
                if n_foto:
                    msg += f" **{n_foto} foto** ikut tersimpan dan akan muncul di file Word."
                else:
                    msg += " (tanpa foto)"
                st.success(msg)

                if ok_drive:
                    st.success("☁️ Laporan Word otomatis tersimpan di Google Drive.")
                    st.link_button("☁️ Buka Arsip di Google Drive", drive_url, use_container_width=True)
                else:
                    st.warning(
                        "Laporan tetap tersimpan di sistem, tetapi arsip Google Drive gagal. "
                        f"{msg_drive}"
                    )

                st.info(
                    "Admin dapat melihat dan menindaklanjuti laporan Anda. "
                    "Cek tab **Laporan Saya** untuk melihat status, membuka arsip Drive, atau mengunduh Word."
                )


def halaman_laporan_saya() -> None:
    st.markdown(
        """
        <div class="main-header">
            <h1>📂 Laporan Saya</h1>
            <p>Daftar laporan yang pernah Anda input — unduh & lihat status</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    username = st.session_state["username"]

    f1, f2 = st.columns(2)
    with f1:
        status_f = st.selectbox(
            "Filter Status", ["Semua"] + STATUS_OPTIONS, key="user_status_f"
        )
    with f2:
        keyword = st.text_input("🔎 Cari cabor / lokasi", key="user_kw")

    df = fetch_laporan_list(
        petugas=username,
        status=status_f if status_f != "Semua" else None,
        keyword=keyword.strip() or None,
    )

    if df.empty:
        st.info(
            "Belum ada laporan. Silakan input melalui tab **Form Input Monitoring**."
        )
        return

    st.caption(f"Total **{len(df)}** laporan Anda.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 📥 Detail & Download")

    pilihan_id = st.selectbox(
        "Pilih laporan",
        df["id"].tolist(),
        format_func=lambda x: (
            f"ID {x} — {df.loc[df['id']==x, 'cabor'].values[0]} "
            f"({df.loc[df['id']==x, 'tanggal'].values[0]}) — "
            f"{df.loc[df['id']==x, 'status'].values[0]}"
        ),
        key="user_pilih_laporan",
    )

    row = get_laporan_by_id(int(pilihan_id))
    if not row:
        st.error("Laporan tidak ditemukan.")
        return

    with st.expander("📋 Lihat Detail Lengkap", expanded=True):
        render_detail_laporan(row)

    st.markdown("### ☁️ Arsip Google Drive")
    if row["drive_url"]:
        st.success(f"{row['drive_status']} — diarsipkan {row['drive_uploaded_at'] or '-'}")
        c1, c2 = st.columns(2)
        with c1:
            st.link_button("☁️ Buka File di Google Drive", row["drive_url"], use_container_width=True)
        with c2:
            if st.button("🔄 Sinkronkan Ulang ke Drive", use_container_width=True, key=f"retry_drive_{row['id']}"):
                ok, msg, _ = retry_drive_upload(int(row["id"]))
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    else:
        st.warning(f"Belum tersimpan di Google Drive: {row['drive_status'] or 'Belum diunggah'}")
        if st.button("☁️ Coba Simpan ke Google Drive", use_container_width=True, key=f"upload_drive_{row['id']}"):
            ok, msg, _ = retry_drive_upload(int(row["id"]))
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    word_file = generate_word_report([dict(row)], is_all=False)
    st.download_button(
        label="⬇️ Download Laporan Word (termasuk foto)",
        data=word_file,
        file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
        mime=WORD_MIME,
        use_container_width=True,
        type="primary",
    )


def halaman_download_laporan() -> None:
    """Halaman khusus user untuk mengunduh file Word dari laporan miliknya."""
    st.markdown(
        """
        <div class="main-header">
            <h1>📥 Download Laporan</h1>
            <p>Unduh file Word dari laporan monitoring yang sudah Anda simpan</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    username = st.session_state["username"]
    status_f = st.selectbox(
        "Filter Status",
        ["Semua"] + STATUS_OPTIONS,
        key="download_status_f",
    )
    keyword = st.text_input(
        "🔎 Cari cabor / lokasi",
        key="download_kw",
        placeholder="Contoh: Tenis Meja / GOR Tangerang",
    )

    df = fetch_laporan_list(
        petugas=username,
        status=status_f if status_f != "Semua" else None,
        keyword=keyword.strip() or None,
    )

    if df.empty:
        st.info("Belum ada laporan yang dapat diunduh.")
        return

    st.caption(f"Ditemukan **{len(df)}** laporan Anda.")

    for _, item in df.iterrows():
        laporan_id = int(item["id"])
        row = get_laporan_by_id(laporan_id)
        if not row:
            continue

        with st.container(border=True):
            c1, c2 = st.columns([4, 1.4])
            with c1:
                st.markdown(
                    f"**ID {laporan_id} — {row['cabor']}**  \n"
                    f"📅 {row['tanggal']} &nbsp; | &nbsp; 📍 {row['lokasi']}  \n"
                    f"📌 Status: **{row['status'] or DEFAULT_STATUS}**",
                    unsafe_allow_html=True,
                )
            with c2:
                word_file = generate_word_report([dict(row)], is_all=False)
                st.download_button(
                    "⬇️ Download Word",
                    data=word_file,
                    file_name=f"Laporan_{row['cabor']}_{row['tanggal']}.docx",
                    mime=WORD_MIME,
                    use_container_width=True,
                    key=f"download_user_word_{laporan_id}",
                )


def halaman_user() -> None:
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_sidebar()

    tab1, tab2, tab3 = st.tabs([
        "📝 Form Input Monitoring",
        "📂 Laporan Saya",
        "📥 Download Laporan",
    ])
    with tab1:
        form_input_monitoring()
    with tab2:
        halaman_laporan_saya()
    with tab3:
        halaman_download_laporan()


# ============================================================
# ENTRY POINT
# ============================================================
init_db()

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    halaman_login()
else:
    if st.session_state.get("role") == "admin":
        halaman_admin()
    else:
        halaman_user()
