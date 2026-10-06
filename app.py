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

def get_foto_bytes(foto: dict) -> Optional[bytes]:
    return foto.get("photo_data")

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
    section.top_margin = Inches(.5); section.bottom_margin = Inches(.5); section.left_margin = Inches(.6); section.right_margin = Inches(.6)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"; style.font.size = Pt(10)
    style.paragraph_format.space_before = Pt(0); style.paragraph_format.space_after = Pt(2); style.paragraph_format.line_spacing = 1.0

    judul = "REKAPITULASI LAPORAN MONEV BINPRES" if is_all else "LAPORAN MONEV BINPRES"
    title = doc.add_heading(judul, level=1); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs: run.font.size = Pt(14); run.font.name = "Calibri"
    meta = doc.add_paragraph(); meta.alignment = WD_ALIGN_PARAGRAPH.CENTER; meta.paragraph_format.space_after = Pt(6)
    _set_run_font(meta.add_run(f"KONI Kabupaten Tangerang  |  Dicetak: {datetime.date.today().strftime('%d/%m/%Y')}"), 9)

    def add_info_line(label: str, value: Any) -> None:
        p = doc.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(1); p.paragraph_format.line_spacing=1.0
        _set_run_font(p.add_run(f"{label}: "),10,True); _set_run_font(p.add_run(str(value) if value else "-"),10)

    def add_section_compact(title_text: str, items: List[Tuple[str,Any]]) -> None:
        h=doc.add_paragraph(); h.paragraph_format.space_before=Pt(6); h.paragraph_format.space_after=Pt(2); _set_run_font(h.add_run(title_text),11,True)
        for question,answer in items:
            pq=doc.add_paragraph(); pq.paragraph_format.space_before=Pt(2); pq.paragraph_format.space_after=Pt(0); pq.paragraph_format.line_spacing=1.0; _set_run_font(pq.add_run(f"• {question}"),9,True)
            pa=doc.add_paragraph(); pa.paragraph_format.space_before=Pt(0); pa.paragraph_format.space_after=Pt(2); pa.paragraph_format.line_spacing=1.0; pa.paragraph_format.left_indent=Inches(.15); _set_run_font(pa.add_run(str(answer).strip() if answer and str(answer).strip() else "—"),9)

    photos_queue=[]
    for idx,data in enumerate(data_list):
        if is_all and idx>0: doc.add_page_break()
        if is_all:
            h=doc.add_paragraph(); h.paragraph_format.space_after=Pt(4); _set_run_font(h.add_run(f"Laporan {idx+1}: {data.get('cabor','-')} — {data.get('tanggal','-')}"),12,True)
        add_info_line("Cabang Olahraga",data.get("cabor")); add_info_line("Tanggal",data.get("tanggal")); add_info_line("Lokasi",data.get("lokasi")); add_info_line("Nama Petugas Monev",data.get("petugas")); add_info_line("Status",data.get("status") or DEFAULT_STATUS); add_info_line("Perwakilan KONI yang Hadir",data.get("daftar_hadir_koni"))
        sep=doc.add_paragraph(); sep.paragraph_format.space_before=Pt(2); sep.paragraph_format.space_after=Pt(2); _set_run_font(sep.add_run("─"*55),8)
        for section_title,fields in FIELD_LABELS.items():
            add_section_compact(section_title,[(label,data.get(key)) for key,label in fields])
        if data.get("catatan_admin"):
            add_section_compact("Catatan Admin / Tindak Lanjut",[("Catatan",data["catatan_admin"])])
        if include_photos and data.get("id"):
            fotos=get_fotos_by_laporan(int(data["id"]),include_data=True)
            if fotos: photos_queue.append((data,fotos))

    if include_photos and photos_queue:
        for data,fotos in photos_queue:
            doc.add_page_break()
            h=doc.add_paragraph(); h.alignment=WD_ALIGN_PARAGRAPH.CENTER; h.paragraph_format.space_after=Pt(4); _set_run_font(h.add_run("DOKUMENTASI FOTO"),12,True)
            sub=doc.add_paragraph(); sub.alignment=WD_ALIGN_PARAGRAPH.CENTER; sub.paragraph_format.space_after=Pt(8); _set_run_font(sub.add_run(f"{data.get('cabor','-')}  |  {data.get('tanggal','-')}  |  {data.get('lokasi','-')}  |  Petugas: {data.get('petugas','-')}"),9)
            for foto in fotos:
                img_bytes=get_foto_bytes(foto)
                if img_bytes:
                    try:
                        doc.add_picture(BytesIO(img_bytes), width=Inches(5.2))
                        cap=doc.add_paragraph(); cap.alignment=WD_ALIGN_PARAGRAPH.CENTER; cap.paragraph_format.space_before=Pt(2); cap.paragraph_format.space_after=Pt(8)
                        _set_run_font(cap.add_run(f"{foto['original_name'] or foto['filename']} | Time-mark: {foto['timestamp_mark'] or '-'}"),8)
                    except Exception: _add_compact_para(doc,f"[Gagal memuat: {foto['original_name']}]",size=9)
                else: _add_compact_para(doc,f"[Foto tidak tersedia: {foto['original_name']}]",size=9)
    buf=BytesIO(); doc.save(buf); return buf.getvalue()

def generate_excel_report(df: pd.DataFrame) -> bytes:
    if df.empty:
        out=BytesIO(); pd.DataFrame({"info":["Tidak ada data"]}).to_excel(out,index=False); return out.getvalue()
    rows=[]
    for laporan_id in [int(x) for x in df["id"].tolist()]:
        row=get_laporan_by_id(laporan_id)
        if row: rows.append(row)
    detail=pd.DataFrame(rows)
    rename_map={
        "id":"ID","tanggal":"Tanggal","cabor":"Cabang Olahraga","lokasi":"Lokasi","petugas":"Nama Petugas",
        "status":"Status","catatan_admin":"Catatan Admin","daftar_hadir_koni":"Daftar Hadir Perwakilan KONI",
        "fisik_parameter":"Fisik - Parameter","fisik_peaking":"Fisik - Peaking","fisik_recovery":"Fisik - Recovery","fisik_cedera":"Fisik - Cedera",
        "taktis_lawan":"Taktis - Lawan","taktis_instruksi":"Taktis - Instruksi","taktis_ujicoba":"Taktis - Ujicoba",
        "mental_cemas":"Mental - Cemas","mental_fokus":"Mental - Fokus","mental_rutinitas":"Mental - Rutinitas","mental_psikolog":"Mental - Psikolog",
        "nutrisi_bb":"Nutrisi - BB","nutrisi_asupan":"Nutrisi - Asupan","nutrisi_hidrasi":"Nutrisi - Hidrasi","nutrisi_tidur":"Nutrisi - Tidur",
        "medis_rekam":"Medis - Rekam","medis_doping":"Medis - Doping","medis_alat":"Medis - Alat","medis_nonteknis":"Medis - Nonteknis","updated_at":"Diperbarui","created_at":"Dibuat"
    }
    detail=detail.rename(columns={k:v for k,v in rename_map.items() if k in detail.columns})
    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer: detail.to_excel(writer,sheet_name="Laporan Monitoring",index=False)
    return out.getvalue()

# ============================================================
# UI COMPONENTS
# ============================================================
def status_badge_html(status: str) -> str:
    s=status or DEFAULT_STATUS
    cls="badge-selesai" if s=="Selesai" else ("badge-sedang" if s=="Sedang Ditindaklanjuti" else "badge-belum")
    return f'<span class="{cls}">{s}</span>'

def render_sidebar() -> None:
    user=get_current_user()
    with st.sidebar:
        st.markdown("## 🏆 BINPRES"); st.caption("Monitoring & Evaluasi Cabor"); st.markdown("---")
        nama=user.get("nama_lengkap") if user and user.get("nama_lengkap") else st.session_state.get("username","")
        st.write(f"👤 **{nama}**"); st.caption(f"Role: {st.session_state.get('role','').upper()}"); st.markdown("---")
        with st.expander("🔑 Ganti Password"):
            with st.form("form_ganti_password"):
                pw_lama=st.text_input("Password lama",type="password"); pw_baru=st.text_input("Password baru",type="password"); pw_konfirm=st.text_input("Konfirmasi password baru",type="password")
                if st.form_submit_button("Simpan Password",use_container_width=True):
                    u=get_current_user()
                    if not pw_lama or not pw_baru: st.error("Semua field wajib diisi.")
                    elif len(pw_baru)<6: st.error("Password baru minimal 6 karakter.")
                    elif pw_baru!=pw_konfirm: st.error("Konfirmasi password tidak cocok.")
                    elif not u or not verify_password(pw_lama,u["password"]): st.error("Password lama salah.")
                    else:
                        conn=db_connect(); conn.execute("UPDATE users SET password=? WHERE id=?",(hash_password(pw_baru),int(u["id"]))); conn.commit(); conn.close(); st.success("Password berhasil diubah.")
        st.markdown("---")
        if st.button("🚪 Logout",use_container_width=True): logout()

def halaman_login() -> None:
    st.markdown(CUSTOM_CSS,unsafe_allow_html=True)
    st.markdown('<div style="text-align:center;padding:40px 0 10px"><div style="font-size:56px">🏆</div><h1 style="margin-bottom:4px;color:#1e3a8a">Sistem Monitoring BINPRES</h1><p style="color:#64748b;font-size:1.05rem">KONI Kabupaten Tangerang</p></div>',unsafe_allow_html=True)
    _,col_c,_=st.columns([1,1.4,1])
    with col_c:
        with st.form("login_form"):
            st.markdown("#### Masuk ke Sistem")
            username=st.text_input("👤 Username",placeholder="Masukkan username")
            password=st.text_input("🔒 Password",type="password",placeholder="Masukkan password")
            submit=st.form_submit_button("Masuk ke Sistem",use_container_width=True,type="primary")
            if submit:
                conn=db_connect(); row=conn.execute("SELECT * FROM users WHERE username=? AND aktif=1 LIMIT 1",(username.strip(),)).fetchone(); conn.close()
                user=dict(row) if row else None
                if user and verify_password(password,user["password"]):
                    st.session_state.update(logged_in=True,username=user["username"],role=user["role"],nama_lengkap=user.get("nama_lengkap") or user["username"]); st.rerun()
                else: st.error("Username atau Password salah / akun tidak aktif.")
        st.caption("Admin: userkoni / koni123 • Petugas: monitoring / koni123")

def render_detail_laporan(row: dict) -> None:
    st.markdown(f"**Cabor:** {row['cabor']} &nbsp;|&nbsp; **Tanggal:** {row['tanggal']} &nbsp;|&nbsp; **Lokasi:** {row['lokasi']} &nbsp;|&nbsp; **Petugas:** {row['petugas']}")
    st.markdown(f"**Status:** {status_badge_html(row.get('status') or DEFAULT_STATUS)}",unsafe_allow_html=True)
    if row.get("catatan_admin"): st.info(f"**Catatan Admin:** {row['catatan_admin']}")
    if row.get("daftar_hadir_koni"):
        st.markdown("#### 👥 Perwakilan KONI Kabupaten Tangerang yang Hadir"); st.write(row["daftar_hadir_koni"])
    for section_title,fields in FIELD_LABELS.items():
        with st.expander(section_title,expanded=False):
            for key,label in fields:
                st.markdown(f"**{label}**"); st.write(row.get(key) or "—")
    fotos=get_fotos_by_laporan(int(row["id"]),include_data=True)
    if fotos:
        st.markdown("#### 🖼️ Dokumentasi Foto")
        cols=st.columns(min(4,len(fotos)))
        for i,foto in enumerate(fotos):
            if foto.get("photo_data"):
                with cols[i%len(cols)]: st.image(foto["photo_data"],caption=foto["original_name"],use_container_width=True)

# ============================================================
# USER / PETUGAS
# ============================================================
def form_input_monitoring() -> None:
    st.markdown('<div class="main-header"><h1>📝 Form Input Monitoring</h1><p>Isi laporan monitoring cabang olahraga secara lengkap</p></div>',unsafe_allow_html=True)
    st.caption(f"Akun: **{st.session_state.get('username','')}**")
    cabor_list=get_cabor_list()
    st.markdown('<span class="section-badge">📷 Dokumentasi Foto</span>',unsafe_allow_html=True)
    st.caption("Upload foto kegiatan sebelum menyimpan form. Maksimal 10 foto, masing-masing maksimal 10 MB.")
    uploaded_files=st.file_uploader("Pilih foto (JPG / PNG / WEBP)",type=["jpg","jpeg","png","webp"],accept_multiple_files=True,key="uploader_monitoring")
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} foto siap disimpan.")
        preview_cols=st.columns(min(4,len(uploaded_files)))
        for i,f in enumerate(uploaded_files):
            with preview_cols[i%len(preview_cols)]: st.image(f,caption=f.name,use_container_width=True)
    else: st.info("Belum ada foto. Foto monitoring wajib minimal 1 foto.")
    st.markdown("---")

    default_petugas=st.session_state.get("nama_lengkap","")
    with st.form("form_monitoring",clear_on_submit=True):
        st.markdown('<span class="section-badge">📌 Informasi Dasar</span>',unsafe_allow_html=True)
        col1,col2=st.columns(2)
        with col1:
            tanggal=st.date_input("Tanggal Monitoring",datetime.date.today())
            cabor=st.selectbox("Cabang Olahraga",["Pilih Cabor..."]+cabor_list)
            nama_petugas=st.text_input("👤 Nama Petugas Monev",value=default_petugas,placeholder="Masukkan nama lengkap petugas")
        with col2:
            lokasi=st.text_input("Lokasi Latihan / Try-out",placeholder="Contoh: GOR Cikokol")
        st.markdown("---")
        with st.expander("💪 1. Performa Fisik & Kebugaran",expanded=True):
            fisik_1=st.text_area("Capaian parameter fisik (vs benchmark target):",height=70); fisik_2=st.text_area("Apakah atlet mencapai grafik performa puncak (peaking)?",height=70); fisik_3=st.text_area("Tingkat pemulihan fisik (recovery):",height=70); fisik_4=st.text_area("Keluhan cedera lama / indikasi cedera baru:",height=70)
        with st.expander("🎯 2. Kesiapan Taktis & Strategi"):
            taktis_1=st.text_area("Pemetaan kekuatan calon lawan:",height=70); taktis_2=st.text_area("Kemampuan mengikuti instruksi teknis di bawah tekanan:",height=70); taktis_3=st.text_area("Hasil try-out / sparing:",height=70)
        with st.expander("🧠 3. Mental, Psikologis & Kesiapan Mental"):
            mental_1=st.text_area("Tingkat kecemasan & kemampuan mengendalikan stres:",height=70); mental_2=st.text_area("Fokus, motivasi, dan self-confidence:",height=70); mental_3=st.text_area("Rutinitas mental khusus saat bertanding:",height=70); mental_4=st.text_area("Koordinasi dengan tim psikolog olahraga:",height=70)
        with st.expander("🥗 4. Nutrisi, Berat Badan & Gaya Hidup"):
            nutrisi_1=st.text_area("Progres penyesuaian berat badan:",height=70); nutrisi_2=st.text_area("Pemantauan asupan nutrisi dan suplemen:",height=70); nutrisi_3=st.text_area("Status hidrasi harian atlet:",height=70); nutrisi_4=st.text_area("Kualitas dan kecukupan waktu tidur:",height=70)
        with st.expander("⚕️ 5. Medis, Bebas Doping & Logistik"):
            medis_1=st.text_area("Status rekam medis terkini & kesiapan fisioterapis:",height=70); medis_2=st.text_area("Keamanan obat, suplemen (Bebas Doping):",height=70); medis_3=st.text_area("Kesiapan perlengkapan khusus bertanding:",height=70); medis_4=st.text_area("Kendala non-teknis (akomodasi, transportasi):",height=70)
        st.markdown("---"); st.markdown('<span class="section-badge">👥 Daftar Hadir Perwakilan KONI Kabupaten Tangerang</span>',unsafe_allow_html=True)
        daftar_hadir_koni=st.text_area("Nama yang hadir mewakili KONI Kabupaten Tangerang",placeholder="Contoh:\n1. Nama — Jabatan\n2. Nama — Jabatan",height=110)
        if uploaded_files: st.caption(f"📷 {len(uploaded_files)} foto siap disimpan. Foto akan diberi time-mark server.")
        else: st.error("📷 Foto monitoring wajib diunggah. Minimal 1 foto.")
        submitted=st.form_submit_button("💾 Simpan Laporan & Siapkan Download Word",use_container_width=True,type="primary")
        if submitted:
            if cabor=="Pilih Cabor...": st.error("⚠️ Harap pilih Cabang Olahraga.")
            elif not lokasi.strip(): st.error("⚠️ Lokasi wajib diisi.")
            elif not nama_petugas.strip(): st.error("⚠️ Nama Petugas wajib diisi.")
            elif not uploaded_files: st.error("⚠️ Foto monitoring wajib diunggah minimal 1 foto.")
            elif len(uploaded_files)>MAX_PHOTOS_PER_REPORT: st.error(f"⚠️ Maksimal {MAX_PHOTOS_PER_REPORT} foto per laporan.")
            elif any(len(f.getbuffer())>MAX_PHOTO_MB*1024*1024 for f in uploaded_files): st.error(f"⚠️ Ukuran setiap foto maksimal {MAX_PHOTO_MB} MB.")
            elif not daftar_hadir_koni.strip(): st.error("⚠️️ Daftar hadir perwakilan KONI wajib diisi.")
            else:
                payload={
                    "tanggal":tanggal.isoformat(),"cabor":cabor,"lokasi":lokasi.strip(),"petugas":nama_petugas.strip(),
                    "fisik_parameter":fisik_1,"fisik_peaking":fisik_2,"fisik_recovery":fisik_3,"fisik_cedera":fisik_4,
                    "taktis_lawan":taktis_1,"taktis_instruksi":taktis_2,"taktis_ujicoba":taktis_3,
                    "mental_cemas":mental_1,"mental_fokus":mental_2,"mental_rutinitas":mental_3,"mental_psikolog":mental_4,
                    "nutrisi_bb":nutrisi_1,"nutrisi_asupan":nutrisi_2,"nutrisi_hidrasi":nutrisi_3,"nutrisi_tidur":nutrisi_4,
                    "medis_rekam":medis_1,"medis_doping":medis_2,"medis_alat":medis_3,"medis_nonteknis":medis_4,
                    "status":DEFAULT_STATUS,"catatan_admin":"","daftar_hadir_koni":daftar_hadir_koni.strip(),"created_at":_now_iso(),"updated_at":_now_iso()
                }
                conn=db_connect()
                cols=", ".join(payload.keys()); placeholders=", ".join(["?"]*len(payload)); cur=conn.execute(f"INSERT INTO laporan_monitoring({cols}) VALUES({placeholders})",tuple(payload.values())); laporan_id=cur.lastrowid; conn.commit(); conn.close()
                n_foto=save_uploaded_photos(int(laporan_id),uploaded_files)
                st.session_state["last_saved_laporan_id"]=int(laporan_id)
                st.success(f"✅ Laporan berhasil disimpan! ID laporan: **{laporan_id}**. {n_foto} foto ikut tersimpan.")
                st.info("File Word dapat langsung diunduh di bawah setelah proses penyimpanan selesai.")

    last_id=st.session_state.get("last_saved_laporan_id")
    if last_id:
        row_last=get_laporan_by_id(int(last_id))
        if row_last:
            st.markdown("---"); st.markdown("### ⬇️ Download Laporan yang Baru Disimpan")
            word_file=generate_word_report([row_last],False,True)
            st.download_button("⬇️ Download Word Laporan Ini (termasuk foto)",data=word_file,file_name=safe_filename(f"Laporan_{row_last['cabor']}_{row_last['tanggal']}.docx"),mime=WORD_MIME,use_container_width=True,type="primary",key=f"dl_after_save_{last_id}")

def halaman_laporan_saya() -> None:
    st.markdown('<div class="main-header"><h1>📂 Laporan Saya</h1><p>Daftar laporan yang pernah Anda input — unduh Word di setiap laporan</p></div>',unsafe_allow_html=True)
    username=st.session_state["username"]
    f1,f2=st.columns(2)
    with f1: status_f=st.selectbox("Filter Status",["Semua"]+STATUS_OPTIONS,key="user_status_f")
    with f2: keyword=st.text_input("🔎 Cari cabor / lokasi / nama petugas",key="user_kw")
    df=fetch_laporan_list(keyword=keyword.strip() or None,status=status_f if status_f!="Semua" else None)
    conn=db_connect(); user=get_current_user(); created_by_name=user.get("nama_lengkap","") if user else ""; conn.close()
    if not df.empty:
        df=df[(df["petugas"].eq(created_by_name)) | (df["petugas"].eq(username))]
    if df.empty: st.info("Belum ada laporan yang sesuai."); return
    st.caption(f"Total **{len(df)}** laporan."); st.dataframe(df,use_container_width=True,hide_index=True)
    st.markdown("---"); st.markdown("### 📥 Detail & Download Word")
    pilihan_id=st.selectbox("Pilih laporan",df["id"].tolist(),format_func=lambda x:f"ID {x} — {df.loc[df['id']==x,'cabor'].values[0]} ({df.loc[df['id']==x,'tanggal'].values[0]}) — {df.loc[df['id']==x,'status'].values[0]}",key="user_pilih_laporan")
    row=get_laporan_by_id(int(pilihan_id))
    if not row: st.error("Laporan tidak ditemukan."); return
    with st.expander("📋 Lihat Detail Lengkap",expanded=True): render_detail_laporan(row)
    st.download_button("⬇️ Download Laporan Word (termasuk foto)",data=generate_word_report([row],False,True),file_name=safe_filename(f"Laporan_{row['cabor']}_{row['tanggal']}.docx"),mime=WORD_MIME,use_container_width=True,type="primary",key=f"user_dl_word_{row['id']}")

def halaman_user() -> None:
    st.markdown(CUSTOM_CSS,unsafe_allow_html=True); render_sidebar()
    tab1,tab2=st.tabs(["📝 Form Input Monitoring","📂 Laporan Saya"])
    with tab1: form_input_monitoring()
    with tab2: halaman_laporan_saya()

# ============================================================
# ADMIN
# ============================================================
def dashboard_admin() -> None:
    stats=fetch_laporan_summary()
    st.markdown('<div class="main-header"><h1>📊 Dashboard Monitoring BINPRES</h1><p>Ringkasan aktivitas monitoring & evaluasi cabang olahraga</p></div>',unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5); c1.metric("Total Laporan",stats["total"]); c2.metric("Bulan Ini",stats["bulan_ini"]); c3.metric("Cabor Termonitor",stats["cabor_termonitor"]); c4.metric("Perlu Tindak Lanjut",stats["perlu_tindak"]); c5.metric("Selesai",stats["selesai"])
    st.markdown("---"); by_cabor,by_month,by_status=fetch_chart_data(); a,b=st.columns(2)
    with a:
        st.markdown("##### 🏅 Top Cabor (jumlah laporan)"); st.bar_chart(by_cabor.set_index("cabor")["jumlah"],height=280) if not by_cabor.empty else st.caption("Belum ada data.")
    with b:
        st.markdown("##### 📅 Tren Laporan per Bulan"); st.line_chart(by_month.set_index("bulan")["jumlah"],height=280) if not by_month.empty else st.caption("Belum ada data.")
    st.markdown("##### 📌 Distribusi Status"); st.bar_chart(by_status.set_index("status")["jumlah"],height=200) if not by_status.empty else st.caption("Belum ada data.")

def kelola_user() -> None:
    st.subheader("👥 Manajemen User"); st.dataframe(fetch_all_users(),use_container_width=True,hide_index=True)
    st.markdown("### ➕ Tambah User")
    with st.form("tambah_user"):
        c1,c2=st.columns(2)
        with c1: username=st.text_input("Username baru"); nama=st.text_input("Nama lengkap")
        with c2: password=st.text_input("Password",type="password"); role=st.selectbox("Hak akses",["user","admin"])
        if st.form_submit_button("Simpan User",use_container_width=True,type="primary"):
            if not username.strip() or not password: st.error("Username dan password wajib diisi.")
            elif len(password)<6: st.error("Password minimal 6 karakter.")
            else:
                try:
                    conn=db_connect(); conn.execute("INSERT INTO users(username,password,role,nama_lengkap,aktif,created_at) VALUES(?,?,?,?,?,?)",(username.strip(),hash_password(password),role,nama.strip(),1,_now_iso())); conn.commit(); conn.close(); st.success("User berhasil ditambahkan."); st.rerun()
                except sqlite3.IntegrityError: st.error("Username sudah digunakan.")
    st.markdown("### ✏️ Edit / Nonaktifkan User")
    conn=db_connect(); users=[dict(r) for r in conn.execute("SELECT id,username,nama_lengkap,role,aktif FROM users ORDER BY username").fetchall()]; conn.close()
    if not users:return
    pilihan=st.selectbox("Pilih user",users,format_func=lambda x:f"{x['username']} — {x.get('nama_lengkap') or '-'}")
    with st.form("edit_user"):
        nama_baru=st.text_input("Nama lengkap",value=pilihan.get("nama_lengkap") or ""); role_baru=st.selectbox("Role",["user","admin"],index=0 if pilihan.get("role")=="user" else 1); aktif_baru=st.checkbox("Akun aktif",value=bool(pilihan.get("aktif"))); password_baru=st.text_input("Password baru (kosongkan jika tidak diubah)",type="password")
        if st.form_submit_button("💾 Simpan Perubahan",use_container_width=True):
            if password_baru and len(password_baru)<6: st.error("Password minimal 6 karakter."); return
            payload=[nama_baru.strip(),role_baru,int(aktif_baru),int(pilihan["id"])]
            conn=db_connect()
            if password_baru: conn.execute("UPDATE users SET nama_lengkap=?,role=?,aktif=?,password=? WHERE id=?",(nama_baru.strip(),role_baru,int(aktif_baru),hash_password(password_baru),int(pilihan["id"])))
            else: conn.execute("UPDATE users SET nama_lengkap=?,role=?,aktif=? WHERE id=?",tuple(payload))
            conn.commit(); conn.close(); st.success("Data user berhasil diperbarui."); st.rerun()

def kelola_cabor() -> None:
    st.subheader("🏅 Kelola Cabang Olahraga"); st.dataframe(fetch_all_cabor(),use_container_width=True,hide_index=True)
    with st.form("tambah_cabor"):
        nama_cabor=st.text_input("Nama Cabang Olahraga Baru")
        if st.form_submit_button("➕ Tambah Cabor",use_container_width=True):
            if not nama_cabor.strip(): st.error("Nama cabor wajib diisi.")
            else:
                try:
                    conn=db_connect(); conn.execute("INSERT INTO cabor(nama,aktif) VALUES(?,1)",(nama_cabor.strip().upper(),)); conn.commit(); conn.close(); st.success("Cabor berhasil ditambahkan."); st.rerun()
                except sqlite3.IntegrityError: st.error("Cabor tersebut sudah ada.")
    st.markdown("### ✏️ Edit / Nonaktifkan Cabor")
    conn=db_connect(); rows=[dict(r) for r in conn.execute("SELECT * FROM cabor ORDER BY nama").fetchall()]; conn.close()
    if not rows:return
    pilihan=st.selectbox("Pilih cabor",rows,format_func=lambda x:x["nama"])
    with st.form("edit_cabor"):
        nama_baru=st.text_input("Nama cabor",value=pilihan["nama"]); aktif=st.checkbox("Cabor aktif",value=bool(pilihan["aktif"]))
        if st.form_submit_button("💾 Simpan",use_container_width=True):
            try:
                conn=db_connect(); conn.execute("UPDATE cabor SET nama=?,aktif=? WHERE id=?",(nama_baru.strip().upper(),int(aktif),int(pilihan["id"]))); conn.commit(); conn.close(); st.success("Cabor berhasil diperbarui."); st.rerun()
            except sqlite3.IntegrityError: st.error("Nama cabor sudah digunakan.")

def delete_laporan(laporan_id:int)->None:
    conn=db_connect(); conn.execute("DELETE FROM laporan_monitoring WHERE id=?",(int(laporan_id),)); conn.commit(); conn.close()

def halaman_laporan_admin() -> None:
    st.subheader("📄 Data Laporan Monitoring")
    f1,f2,f3,f4=st.columns(4)
    with f1: cabor_filter=st.selectbox("Filter Cabor",["Semua"]+get_cabor_list())
    with f2: status_filter=st.selectbox("Filter Status",["Semua"]+STATUS_OPTIONS)
    with f3: tgl_awal=st.date_input("Dari tanggal",value=None)
    with f4: tgl_akhir=st.date_input("Sampai tanggal",value=None)
    keyword=st.text_input("🔎 Cari lokasi / petugas / cabor",placeholder="Ketik kata kunci...")
    df=fetch_laporan_list(cabor=cabor_filter,status=status_filter,tgl_awal=tgl_awal,tgl_akhir=tgl_akhir,keyword=keyword.strip() or None)
    st.caption(f"Menampilkan **{len(df)}** laporan.")
    if df.empty: st.info("Tidak ada laporan sesuai filter."); return
    st.dataframe(df,use_container_width=True,hide_index=True); st.markdown("---"); st.markdown("### 🔍 Detail, Export & Tindak Lanjut")
    pilihan_id=st.selectbox("Pilih laporan",df["id"].tolist(),format_func=lambda x:f"ID {x} — {df.loc[df['id']==x,'cabor'].values[0]} ({df.loc[df['id']==x,'tanggal'].values[0]}) — {df.loc[df['id']==x,'status'].values[0]}",key="admin_pilih_laporan")
    row=get_laporan_by_id(int(pilihan_id))
    if not row: st.error("Laporan tidak ditemukan."); return
    with st.expander("📋 Lihat Detail Lengkap",expanded=False): render_detail_laporan(row)
    col_w,col_x,col_d=st.columns(3)
    with col_w:
        st.download_button("⬇️ Word Satuan",data=generate_word_report([row],False,True),file_name=safe_filename(f"Laporan_{row['cabor']}_{row['tanggal']}.docx"),mime=WORD_MIME,use_container_width=True)
    with col_x:
        st.download_button("⬇️ Excel (hasil filter)",data=generate_excel_report(df),file_name=f"Rekap_Monev_{datetime.date.today()}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    with col_d:
        if st.button("🗑️ Hapus Laporan",use_container_width=True): st.session_state["confirm_delete_id"]=int(pilihan_id)
    if st.session_state.get("confirm_delete_id")==int(pilihan_id):
        st.warning(f"Yakin hapus laporan ID {pilihan_id}? Tindakan ini tidak dapat dibatalkan.")
        y,n=st.columns(2)
        with y:
            if st.button("✅ Ya, Hapus",type="primary",use_container_width=True): delete_laporan(int(pilihan_id)); st.session_state.pop("confirm_delete_id",None); st.success("Laporan berhasil dihapus."); st.rerun()
        with n:
            if st.button("❌ Batal",use_container_width=True): st.session_state.pop("confirm_delete_id",None); st.rerun()
    st.markdown("#### 📚 Export Rekap Word (semua hasil filter)")
    rows=[get_laporan_by_id(int(x)) for x in df["id"].tolist()]; rows=[r for r in rows if r]
    st.download_button("⬇️ Download Rekap Word",data=generate_word_report(rows,True,True),file_name=f"Rekap_Monev_{datetime.date.today()}.docx",mime=WORD_MIME,use_container_width=True)
    st.markdown("---"); st.markdown("### ✏️ Tindak Lanjut Laporan")
    with st.form("update_status_laporan"):
        current=row.get("status") or DEFAULT_STATUS; status=st.selectbox("Status tindak lanjut",STATUS_OPTIONS,index=STATUS_OPTIONS.index(current) if current in STATUS_OPTIONS else 0); catatan=st.text_area("Catatan Admin / Rekomendasi",value=row.get("catatan_admin") or "",height=100)
        if st.form_submit_button("💾 Simpan Tindak Lanjut",use_container_width=True):
            conn=db_connect(); conn.execute("UPDATE laporan_monitoring SET status=?,catatan_admin=?,updated_at=? WHERE id=?",(status,catatan,_now_iso(),int(pilihan_id))); conn.commit(); conn.close(); st.success("Tindak lanjut berhasil diperbarui."); st.rerun()

def halaman_admin() -> None:
    st.markdown(CUSTOM_CSS,unsafe_allow_html=True); render_sidebar(); dashboard_admin()
    tab1,tab2,tab3=st.tabs(["📄 Laporan Monitoring","👥 Manajemen User","🏅 Kelola Cabor"])
    with tab1: halaman_laporan_admin()
    with tab2: kelola_user()
    with tab3: kelola_cabor()

# ============================================================
# ENTRY POINT
# ============================================================
init_backend()
if "logged_in" not in st.session_state: st.session_state["logged_in"]=False
if not st.session_state["logged_in"]: halaman_login()
elif st.session_state.get("role")=="admin": halaman_admin()
else: halaman_user()
