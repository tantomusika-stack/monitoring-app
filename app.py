from __future__ import annotations

import streamlit as st
import datetime
import hashlib
import sqlite3
import re
from io import BytesIO
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# KONFIGURASI
# ============================================================
st.set_page_config(
    page_title="Sistem Monitoring BINPRES",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Database akan tersimpan dengan aman di folder yang sama dengan app.py
DB_PATH = Path("monitoring_binpres.db")
WORD_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_PHOTO_MB = 10
MAX_PHOTOS_PER_REPORT = 10
PHOTO_TIMESTAMP_FORMAT = "%d-%m-%Y %H:%M:%S"
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
.main-header{background:linear-gradient(135deg,#1e3a8a 0%,#3b82f6 100%);color:white;padding:1.2rem 1.5rem;border-radius:12px;margin-bottom:1.5rem;box-shadow:0 4px 12px rgba(30,58,138,.25)}
.main-header h1{margin:0;font-size:1.6rem;font-weight:700}.main-header p{margin:.3rem 0 0;opacity:.9;font-size:.95rem}
div[data-testid="stMetric"]{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:.8rem 1rem;box-shadow:0 1px 3px rgba(0,0,0,.06)}
div[data-testid="stMetric"] label{color:#64748b!important;font-size:.85rem!important}div[data-testid="stMetric"] [data-testid="stMetricValue"]{color:#1e3a8a!important;font-weight:700}
.section-badge{display:inline-block;background:#1e3a8a;color:white;padding:.35rem .9rem;border-radius:6px;font-weight:600;font-size:.95rem;margin-bottom:.8rem}
.badge-belum{background:#fef3c7;color:#92400e;padding:3px 10px;border-radius:12px;font-size:.8rem;font-weight:600}.badge-sedang{background:#dbeafe;color:#1e40af;padding:3px 10px;border-radius:12px;font-size:.8rem;font-weight:600}.badge-selesai{background:#d1fae5;color:#065f46;padding:3px 10px;border-radius:12px;font-size:.8rem;font-weight:600}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#0f172a 0%,#1e293b 100%)}section[data-testid="stSidebar"] *{color:#e2e8f0!important}
section[data-testid="stSidebar"] .stButton>button{background:#334155;border:1px solid #475569;color:white}section[data-testid="stSidebar"] .stButton>button:hover{background:#475569;border-color:#64748b}
.stButton>button[kind="primary"],div[data-testid="stFormSubmitButton"]>button{background:linear-gradient(135deg,#dc2626 0%,#ef4444 100%)!important;border:none!important;font-weight:600!important;border-radius:8px!important}
div[data-testid="stExpander"]{border:1px solid #e2e8f0;border-radius:8px;margin-bottom:.5rem}#MainMenu{visibility:hidden}footer{visibility:hidden}
</style>
"""

# ============================================================
# SQLITE BACKEND
# ============================================================
def db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def verify_password(password: str, stored: str) -> bool:
    return hash_password(password) == stored or password == stored

def _now_iso() -> str:
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")

def safe_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r'[\\/:*?"<>|]+', "-", value)
    value = re.sub(r"\s+", " ", value)
    return value[:180] or "Laporan_Monitoring"

def init_backend() -> None:
    conn = db_connect()
    try:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user',
            nama_lengkap TEXT NOT NULL DEFAULT '',
            aktif INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS cabor (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT NOT NULL UNIQUE,
            aktif INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS laporan_monitoring (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT NOT NULL,
            cabor TEXT NOT NULL,
            lokasi TEXT NOT NULL,
            petugas TEXT NOT NULL,
            fisik_parameter TEXT DEFAULT '', fisik_peaking TEXT DEFAULT '', fisik_recovery TEXT DEFAULT '', fisik_cedera TEXT DEFAULT '',
            taktis_lawan TEXT DEFAULT '', taktis_instruksi TEXT DEFAULT '', taktis_ujicoba TEXT DEFAULT '',
            mental_cemas TEXT DEFAULT '', mental_fokus TEXT DEFAULT '', mental_rutinitas TEXT DEFAULT '', mental_psikolog TEXT DEFAULT '',
            nutrisi_bb TEXT DEFAULT '', nutrisi_asupan TEXT DEFAULT '', nutrisi_hidrasi TEXT DEFAULT '', nutrisi_tidur TEXT DEFAULT '',
            medis_rekam TEXT DEFAULT '', medis_doping TEXT DEFAULT '', medis_alat TEXT DEFAULT '', medis_nonteknis TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'Belum Ditindaklanjuti',
            catatan_admin TEXT DEFAULT '',
            daftar_hadir_koni TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS laporan_foto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            laporan_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            original_name TEXT NOT NULL,
            mime_type TEXT NOT NULL,
            photo_data BLOB NOT NULL,
            timestamp_mark TEXT NOT NULL,
            FOREIGN KEY(laporan_id) REFERENCES laporan_monitoring(id) ON DELETE CASCADE
        );
        """)
        now = _now_iso()
        # Create Default Admin & User
        conn.execute("INSERT OR IGNORE INTO users(username,password,role,nama_lengkap,aktif,created_at) VALUES(?,?,?,?,?,?)",
            ("userkoni", hash_password("koni123"), "admin", "Admin KONI", 1, now))
        conn.execute("INSERT OR IGNORE INTO users(username,password,role,nama_lengkap,aktif,created_at) VALUES(?,?,?,?,?,?)",
            ("monitoring", hash_password("koni123"), "user", "Petugas Monitoring", 1, now))
        for nama in DEFAULT_CABOR:
            conn.execute("INSERT OR IGNORE INTO cabor(nama,aktif) VALUES(?,1)", (nama,))
        conn.commit()
    finally:
        conn.close()

def logout() -> None:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# ============================================================
# QUERY / HELPER
# ============================================================
def get_cabor_list() -> List[str]:
    conn = db_connect()
    rows = conn.execute("SELECT nama FROM cabor WHERE aktif=1 ORDER BY nama").fetchall()
    conn.close()
    return [r["nama"] for r in rows]

def get_current_user() -> Optional[dict]:
    username = st.session_state.get("username")
    if not username: return None
    conn = db_connect()
    row = conn.execute("SELECT * FROM users WHERE username=? LIMIT 1", (username,)).fetchone()
    conn.close()
    return dict(row) if row else None

def fetch_all_users() -> pd.DataFrame:
    conn = db_connect()
    df = pd.read_sql_query("SELECT id,username,nama_lengkap,role,aktif,created_at FROM users ORDER BY id DESC", conn)
    conn.close()
    if df.empty: return pd.DataFrame(columns=["id","username","nama_lengkap","role","status","created_at"])
    df["status"] = df["aktif"].map({1:"Aktif",0:"Nonaktif"})
    return df[["id","username","nama_lengkap","role","status","created_at"]]

def fetch_all_cabor() -> pd.DataFrame:
    conn = db_connect()
    df = pd.read_sql_query("SELECT id,nama,aktif FROM cabor ORDER BY nama", conn)
    conn.close()
    if df.empty: return pd.DataFrame(columns=["id","nama","status"])
    df["status"] = df["aktif"].map({1:"Aktif",0:"Nonaktif"})
    return df[["id","nama","status"]]

def _all_laporan_rows() -> List[dict]:
    conn = db_connect()
    rows = conn.execute("SELECT * FROM laporan_monitoring ORDER BY tanggal DESC,id DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def fetch_laporan_summary() -> Dict[str, int]:
    rows = _all_laporan_rows()
    bulan_ini = datetime.date.today().strftime("%Y-%m")
    return {
        "total": len(rows),
        "bulan_ini": sum(str(r.get("tanggal", ""))[:7] == bulan_ini for r in rows),
        "cabor_termonitor": len({r.get("cabor") for r in rows if r.get("cabor")}),
        "perlu_tindak": sum((r.get("status") or DEFAULT_STATUS) == DEFAULT_STATUS for r in rows),
        "selesai": sum((r.get("status") or "") == "Selesai" for r in rows),
    }

def fetch_chart_data():
    rows = _all_laporan_rows()
    if not rows:
        return (pd.DataFrame(columns=["cabor","jumlah"]), pd.DataFrame(columns=["bulan","jumlah"]), pd.DataFrame(columns=["status","jumlah"]))
    df = pd.DataFrame(rows)
    by_cabor = df.groupby("cabor").size().reset_index(name="jumlah").sort_values("jumlah", ascending=False).head(15)
    df["bulan"] = df["tanggal"].astype(str).str[:7]
    by_month = df.groupby("bulan").size().reset_index(name="jumlah").sort_values("bulan")
    df["status"] = df["status"].fillna(DEFAULT_STATUS)
    by_status = df.groupby("status").size().reset_index(name="jumlah")
    return by_cabor, by_month, by_status

def fetch_laporan_list(petugas=None, cabor=None, status=None, tgl_awal=None, tgl_akhir=None, keyword=None) -> pd.DataFrame:
    rows = _all_laporan_rows()
    columns = ["id","tanggal","cabor","lokasi","petugas","status","updated_at"]
    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame(columns=columns)
    if petugas: df = df[df["petugas"].fillna("").eq(petugas)]
    if cabor and cabor != "Semua": df = df[df["cabor"].fillna("").eq(cabor)]
    if status and status != "Semua": df = df[df["status"].fillna(DEFAULT_STATUS).eq(status)]
    if tgl_awal: df = df[df["tanggal"].astype(str).ge(str(tgl_awal))]
    if tgl_akhir: df = df[df["tanggal"].astype(str).le(str(tgl_akhir))]
    if keyword:
        kw = keyword.lower()
        mask = (df["lokasi"].fillna("").str.lower().str.contains(kw, regex=False) |
                df["petugas"].fillna("").str.lower().str.contains(kw, regex=False) |
                df["cabor"].fillna("").str.lower().str.contains(kw, regex=False))
        df = df[mask]
    df["status"] = df["status"].fillna(DEFAULT_STATUS)
    return df[columns].sort_values(["tanggal","id"], ascending=[False,False]).reset_index(drop=True)

def get_laporan_by_id(laporan_id: int) -> Optional[dict]:
    conn = db_connect()
    row = conn.execute("SELECT * FROM laporan_monitoring WHERE id=? LIMIT 1", (int(laporan_id),)).fetchone()
    conn.close()
    return dict(row) if row else None

def get_fotos_by_laporan(laporan_id: int, include_data: bool = True) -> List[dict]:
    conn = db_connect()
    cols = "id,laporan_id,filename,original_name,mime_type,photo_data,timestamp_mark" if include_data else "id,laporan_id,filename,original_name,mime_type,timestamp_mark"
    rows = conn.execute(f"SELECT {cols} FROM laporan_foto WHERE laporan_id=? ORDER BY id", (int(laporan_id),)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

# ============================================================
# FOTO & WORD
# ============================================================
def _load_font(size: int = 24):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/System/Library/Fonts/Helvetica.ttc"
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            continue
    return ImageFont.load_default()

def add_timestamp_watermark(raw: bytes, timestamp_text: str) -> bytes:
    try:
        img = Image.open(BytesIO(raw)).convert("RGB")
        max_side = 2400
        if max(img.size) > max_side: img.thumbnail((max_side,max_side), Image.Resampling.LANCZOS)
        draw = ImageDraw.Draw(img, "RGBA")
        font_size = max(18, int(min(img.size) * 0.025))
        font = _load_font(font_size)
        text = f"MONITORING BINPRES | {timestamp_text}"
        bbox = draw.textbbox((0,0), text, font=font)
        tw,th = bbox[2]-bbox[0], bbox[3]-bbox[1]
        margin = max(12,int(font_size*.6))
        x, y = margin, img.height-th-margin*2
        draw.rounded_rectangle((x-margin,y-margin,x+tw+margin,y+th+margin),radius=10,fill=(0,0,0,155))
        draw.text((x,y), text, font=font, fill=(255,255,255,235))
        out = BytesIO()
        img.save(out, format="JPEG", quality=90, optimize=True)
        return out.getvalue()
    except Exception: return raw

def save_uploaded_photos(laporan_id: int, uploaded_files: list) -> int:
    if not uploaded_files: return 0
    conn = db_connect()
    count = 0
    for f in uploaded_files[:MAX_PHOTOS_PER_REPORT]:
        if Path(f.name).suffix.lower() not in {".jpg",".jpeg",".png",".webp"}: continue
        raw = f.getbuffer().tobytes()
        if not raw: continue
        timestamp_text = datetime.datetime.now().astimezone().strftime(PHOTO_TIMESTAMP_FORMAT)
        processed = add_timestamp_watermark(raw, timestamp_text)
        filename = f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{safe_filename(Path(f.name).stem)}.jpg"
        conn.execute("INSERT INTO laporan_foto(laporan_id,filename,original_name,mime_type,photo_data,timestamp_mark) VALUES(?,?,?,?,?,?)",
                     (laporan_id, filename, f.name, "image/jpeg", sqlite3.Binary(processed), timestamp_text))
        count += 1
    conn.commit()
    conn.close()
    return count

def _set_run_font(run, size_pt: float = 10, bold: bool = False) -> None:
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.name = "Calibri"

def _add_compact_para(doc, text: str, bold: bool = False, size: float = 10, space_after: float = 2) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after, p.paragraph_format.line_spacing = Pt(0), Pt(space_after), 1.0
    _set_run_font(p.add_run(text), size_pt=size, bold=bold)

def generate_word_report(data_list: List[Dict[str, Any]], is_all: bool = False, include_photos: bool = True) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(.5)
    section.left_margin = section.right_margin = Inches(.6)
    style = doc.styles["Normal"]
    style.font.name, style.font.size = "Calibri", Pt(10)
    
    title = doc.add_heading("REKAPITULASI LAPORAN MONEV BINPRES" if is_all else "LAPORAN MONEV BINPRES", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs: run.font.size, run.font.name = Pt(14), "Calibri"
    
    meta = doc.add_paragraph(); meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_run_font(meta.add_run(f"KONI Kabupaten Tangerang  |  Dicetak: {datetime.date.today().strftime('%d/%m/%Y')}"), 9)

    def add_info(label, value): _add_compact_para(doc, f"{label}: {value or '-'}", bold=True)
    def add_section(title_text, items):
        doc.add_paragraph().add_run(title_text).bold = True
        for q, a in items: _add_compact_para(doc, f"• {q}\n  {str(a).strip() if a else '—'}")

    for idx, data in enumerate(data_list):
        if is_all and idx > 0: doc.add_page_break()
        add_info("Cabang Olahraga", data.get("cabor"))
        add_info("Tanggal & Lokasi", f"{data.get('tanggal')} di {data.get('lokasi')}")
        add_info("Petugas", data.get("petugas"))
        
        for sec_title, fields in FIELD_LABELS.items():
            add_section(sec_title, [(label, data.get(k)) for k, label in fields])
            
        if include_photos and data.get("id"):
            fotos = get_fotos_by_laporan(int(data["id"]))
            if fotos:
                doc.add_page_break()
                doc.add_heading("DOKUMENTASI FOTO", level=2)
                for foto in fotos:
                    if foto.get("photo_data"):
                        try:
                            doc.add_picture(BytesIO(foto["photo_data"]), width=Inches(5.0))
                            _add_compact_para(doc, f"{foto['original_name']} | Time-mark: {foto['timestamp_mark']}", size=8)
                        except: pass
    buf = BytesIO(); doc.save(buf); return buf.getvalue()

def generate_excel_report(df: pd.DataFrame) -> bytes:
    if df.empty:
        out=BytesIO(); pd.DataFrame({"info":["Tidak ada data"]}).to_excel(out,index=False); return out.getvalue()
    rows = [get_laporan_by_id(int(x)) for x in df["id"].tolist() if get_laporan_by_id(int(x))]
    detail = pd.DataFrame(rows)
    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer: detail.to_excel(writer,sheet_name="Laporan Monitoring",index=False)
    return out.getvalue()

# ============================================================
# UI COMPONENTS
# ============================================================
def render_sidebar():
    user = get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES")
        st.write(f"👤 **{user.get('nama_lengkap', st.session_state.get('username')) if user else ''}**")
        if st.button("🚪 Logout", use_container_width=True): logout()

def halaman_login():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.4, 1])
    with col:
        with st.form("login"):
            st.markdown("### Sistem Monitoring BINPRES")
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Masuk", type="primary"):
                conn = db_connect()
                user = conn.execute("SELECT * FROM users WHERE username=? AND aktif=1", (u,)).fetchone()
                conn.close()
                if user and verify_password(p, user["password"]):
                    st.session_state.update(logged_in=True, username=user["username"], role=user["role"])
                    st.rerun()
                else: st.error("Login Gagal")

def halaman_admin():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True); render_sidebar()
    st.title("Admin Dashboard")
    df = fetch_laporan_list()
    st.dataframe(df, use_container_width=True)
    if not df.empty:
        pilihan = st.selectbox("Pilih ID Laporan untuk diunduh", df["id"].tolist())
        row = get_laporan_by_id(int(pilihan))
        if row:
            st.download_button("⬇️ Download Word", data=generate_word_report([row]), file_name=f"Laporan_{row['cabor']}.docx", mime=WORD_MIME)

def halaman_user():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True); render_sidebar()
    st.title("Form Monitoring")
    with st.form("input"):
        cabor = st.selectbox("Cabor", ["Pilih..."] + get_cabor_list())
        lokasi = st.text_input("Lokasi")
        petugas = st.text_input("Petugas", value=st.session_state.get('username'))
        fotos = st.file_uploader("Upload Foto", accept_multiple_files=True, type=["jpg","png"])
        if st.form_submit_button("Simpan Laporan", type="primary"):
            if cabor != "Pilih..." and lokasi:
                conn = db_connect()
                cur = conn.execute("INSERT INTO laporan_monitoring (tanggal,cabor,lokasi,petugas,created_at,updated_at) VALUES (?,?,?,?,?,?)",
                                   (_now_iso()[:10], cabor, lokasi, petugas, _now_iso(), _now_iso()))
                laporan_id = cur.lastrowid
                conn.commit(); conn.close()
                save_uploaded_photos(laporan_id, fotos)
                st.success("Laporan Tersimpan!")
            else: st.error("Lengkapi data dasar!")

# ============================================================
# ENTRY POINT
# ============================================================
init_backend()
if not st.session_state.get("logged_in"): halaman_login()
elif st.session_state.get("role") == "admin": halaman_admin()
else: halaman_user()
