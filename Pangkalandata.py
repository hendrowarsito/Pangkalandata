import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import folium
from folium.plugins import HeatMap
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
import numpy as np

st.set_page_config(
    layout="wide",
    initial_sidebar_state="expanded",
    page_title="Pangkalan Data Tanah KJPP SRR",
    page_icon=":material/landscape:"
)

# ─── Palet & gaya global ─────────────────────────────────────────────────────
NAVY    = "#1F4E79"
NAVY_DK = "#0F2A44"
TEAL    = "#2A9D8F"
AMBER   = "#E9A23B"
RED     = "#C8553D"
SLATE   = "#64748B"
LINE    = "#E2E8F0"

st.markdown(f"""
<style>
.block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1500px; }}
h1, h2, h3, h4 {{ letter-spacing: -0.01em; }}
[data-testid="stSidebar"] {{ border-right: 1px solid {LINE}; }}
[data-testid="stSidebar"] .block-container {{ padding-top: 1.5rem; }}

/* Header aplikasi */
.app-header {{ display:flex; align-items:flex-end; justify-content:space-between;
               gap:1rem; flex-wrap:wrap; padding-bottom:.9rem; margin-bottom:.4rem;
               border-bottom:1px solid {LINE}; }}
.app-header h1 {{ font-size:1.55rem; font-weight:700; color:{NAVY_DK}; margin:0; padding:0; }}
.app-header .sub {{ color:{SLATE}; font-size:.9rem; margin-top:.15rem; }}
.chips {{ display:flex; gap:.4rem; flex-wrap:wrap; }}
.chip {{ background:#F1F5F9; border:1px solid {LINE}; color:#334155; border-radius:999px;
         padding:.18rem .7rem; font-size:.78rem; white-space:nowrap; }}
.chip b {{ color:{NAVY_DK}; font-weight:600; }}

/* Kartu KPI */
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr));
             gap:.75rem; margin:.4rem 0 1rem; }}
.kpi {{ background:#fff; border:1px solid {LINE}; border-radius:10px; padding:.85rem 1rem; }}
.kpi .lbl {{ color:{SLATE}; font-size:.72rem; font-weight:600; text-transform:uppercase;
             letter-spacing:.05em; }}
.kpi .val {{ color:{NAVY_DK}; font-size:1.3rem; font-weight:700; margin-top:.25rem;
             font-variant-numeric:tabular-nums; white-space:nowrap; }}
.kpi .val small {{ font-size:.75rem; font-weight:500; color:{SLATE}; margin-left:.15rem; }}
.kpi .note {{ color:{SLATE}; font-size:.75rem; margin-top:.15rem; }}
.kpi.accent {{ background:{NAVY}; border-color:{NAVY}; }}
.kpi.accent .lbl, .kpi.accent .note, .kpi.accent .val small {{ color:#C9D8E8; }}
.kpi.accent .val {{ color:#fff; }}

/* Badge status */
.badge {{ display:inline-block; border-radius:999px; padding:.12rem .6rem; font-size:.72rem;
          font-weight:600; }}
.badge.ok   {{ background:#E6F4F1; color:#1E7A6F; }}
.badge.warn {{ background:#FDF3E1; color:#9A6412; }}
.badge.bad  {{ background:#FBEAE6; color:#A23F2B; }}

/* Judul bagian & langkah */
.section-title {{ font-size:1rem; font-weight:650; color:{NAVY_DK}; margin:0 0 .5rem; }}
.step {{ display:inline-flex; align-items:center; justify-content:center; width:1.4rem;
         height:1.4rem; border-radius:50%; background:{NAVY}; color:#fff; font-size:.75rem;
         font-weight:700; margin-right:.45rem; }}

/* Daftar label–nilai */
.kv {{ display:grid; grid-template-columns:auto 1fr; gap:.3rem .9rem; font-size:.86rem; }}
.kv dt {{ color:{SLATE}; }}
.kv dd {{ margin:0; color:#1E293B; font-weight:500; }}

/* Sidebar */
.brand {{ display:flex; gap:.65rem; align-items:center; margin-bottom:.4rem; }}
.brand .logo {{ width:38px; height:38px; border-radius:9px; background:{NAVY}; color:#fff;
                display:flex; align-items:center; justify-content:center; font-weight:800;
                font-size:.9rem; letter-spacing:.02em; }}
.brand .t1 {{ font-weight:700; color:{NAVY_DK}; font-size:.98rem; line-height:1.15; }}
.brand .t2 {{ color:{SLATE}; font-size:.74rem; }}
.side-label {{ color:{SLATE}; font-size:.72rem; font-weight:700; text-transform:uppercase;
               letter-spacing:.06em; margin:1.1rem 0 .4rem; }}
.status-line {{ font-size:.8rem; color:#334155; display:flex; gap:.45rem; align-items:baseline;
                margin:.2rem 0; }}
.dot {{ width:.5rem; height:.5rem; border-radius:50%; display:inline-block; flex-shrink:0; }}

/* Halaman awal */
.hero {{ padding:2.2rem 0 1.2rem; }}
.hero .eyebrow {{ color:{TEAL}; font-weight:700; font-size:.78rem; letter-spacing:.08em;
                  text-transform:uppercase; }}
.hero h1 {{ font-size:2.1rem; color:{NAVY_DK}; margin:.3rem 0 .4rem; font-weight:750; }}
.hero p {{ color:{SLATE}; font-size:1.02rem; max-width:680px; }}
.feat-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:.9rem;
              margin:1rem 0 1.5rem; }}
.feat {{ border:1px solid {LINE}; border-radius:12px; padding:1.1rem 1.15rem; background:#fff; }}
.feat .n {{ color:{TEAL}; font-weight:700; font-size:.8rem; }}
.feat h4 {{ margin:.3rem 0 .3rem; font-size:1rem; color:{NAVY_DK}; padding:0; }}
.feat p {{ margin:0; color:{SLATE}; font-size:.86rem; line-height:1.45; }}

/* Tab */
.stTabs [data-baseweb="tab-list"] {{ gap:1.4rem; border-bottom:1px solid {LINE}; }}
.stTabs [data-baseweb="tab"] {{ padding:.6rem 0; font-weight:600; }}
</style>
""", unsafe_allow_html=True)

YEAR_COLORS = {
    "gte_2025": TEAL,
    "gte_2024": NAVY,
    "gte_2023": AMBER,
    "lt_2023":  RED,
    "subject":  NAVY_DK,
}

PLOT_FONT = dict(family="Source Sans Pro, Segoe UI, sans-serif", size=12, color="#334155")
PLOT_CFG  = {"displayModeBar": False}
PLOT_SEQ  = [NAVY, TEAL, AMBER, RED, "#7A8CA5", "#8E6C8A", "#5B8C5A"]

def style_fig(fig, title=None, height=340):
    """Gaya grafik seragam: latar putih, grid tipis, judul rata kiri."""
    fig.update_layout(
        title=dict(text=title, x=0, xanchor="left", y=0.97, yref="container", yanchor="top",
                   font=dict(size=14, color=NAVY_DK)) if title else None,
        template="simple_white", font=PLOT_FONT, height=height,
        margin=dict(t=76 if title else 24, b=8, l=8, r=8),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(bgcolor="white", font_size=12),
        colorway=PLOT_SEQ,
    )
    fig.update_xaxes(showgrid=False, linecolor=LINE, ticks="", title_font=dict(size=12, color=SLATE))
    fig.update_yaxes(showgrid=True, gridcolor="#EEF2F6", linecolor=LINE, ticks="",
                     title_font=dict(size=12, color=SLATE), separatethousands=True)
    return fig

def kpi_cards(items):
    """items: list of (label, value, note, accent) → baris kartu KPI."""
    html = "".join(
        f'<div class="kpi{" accent" if accent else ""}"><div class="lbl">{lbl}</div>'
        f'<div class="val">{val}</div>'
        + (f'<div class="note">{note}</div>' if note else "") + "</div>"
        for lbl, val, note, accent in items
    )
    st.markdown(f'<div class="kpi-grid">{html}</div>', unsafe_allow_html=True)

def section_title(text, step=None):
    num = f'<span class="step">{step}</span>' if step else ""
    st.markdown(f'<div class="section-title">{num}{text}</div>', unsafe_allow_html=True)

def rp(value, unit=True):
    """Format Rupiah untuk kartu: 'Rp 1.234.567' dengan satuan /m² kecil."""
    try:
        s = f"Rp {float(value):,.0f}".replace(",", ".")
    except Exception:
        return "—"
    return s + ("<small>/m²</small>" if unit else "")

# Label tampilan untuk kolom internal
COL_LABELS = {
    "Nomor": "No. Data", "Jenis_Data": "Jenis Data", "Tanggal_Data": "Tgl. Data",
    "Harga_Total": "Harga Total (Rp)", "Harga_Tanah": "Harga Tanah (Rp/m²)",
    "Luas_Tanah": "Luas Tanah (m²)", "Luas_Bangunan": "Luas Bangunan (m²)",
    "Kondisi_Bangunan": "Kondisi Bangunan (%)", "Kelas_Bangunan": "Kelas Bangunan",
    "Jenis_Properti": "Jenis Properti", "Sumber_Data": "Sumber Data", "Telp": "Telepon",
    "Telp_Konfirmasi": "Telp. Konfirmasi", "Kota": "Kabupaten/Kota", "Kelurahan": "Desa/Kelurahan",
    "Kode_Inspeksi": "Kode Inspeksi", "Tanggal_Inspeksi": "Tgl. Inspeksi",
    "Pemberi_Tugas": "Pemberi Tugas", "Foto_Jalan": "Foto Jalan", "Foto_Dalam": "Foto Dalam",
    "Foto_Samping_Kanan": "Foto Samping Kanan", "Foto_Samping_Kiri": "Foto Samping Kiri",
    "Gambar_Situasi": "Gambar Situasi", "Tahun_Bersih": "Tahun",
}

def generate_streetview_url(lat, lon):
    return f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}&heading=0&pitch=0&fov=75"

def gdrive_file_id(url):
    """Ekstrak file ID dari berbagai format URL Google Drive."""
    if not url or str(url).strip() in ("", "#", "nan", "None", "-"):
        return None
    url = str(url).strip()
    patterns = [
        r"drive\.google\.com/file/d/([a-zA-Z0-9_-]+)",
        r"drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)",
        r"drive\.google\.com/uc\?(?:[^&]*&)*id=([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
    ]
    for pat in patterns:
        m = re.search(pat, url)
        if m:
            return m.group(1)
    return None

def gdrive_thumbnail(url, width=300):
    """
    Coba beberapa format URL thumbnail Google Drive.
    Kembalikan daftar URL (diurutkan dari paling reliable).
    """
    fid = gdrive_file_id(url)
    if not fid:
        return None, None
    # lh3 = CDN langsung, tidak perlu redirect auth (lebih andal di iframe)
    lh3  = f"https://lh3.googleusercontent.com/d/{fid}=w{width}"
    # thumbnail = API resmi Google Drive (backup)
    thumb = f"https://drive.google.com/thumbnail?id={fid}&sz=w{width}"
    return lh3, thumb

def format_currency(value):
    try:
        return f"Rp {float(value):,.0f}".replace(",", ".")
    except Exception:
        return "N/A"

def bersihkan_tahun(val):
    try:
        return pd.to_numeric(str(val).replace(",", "").strip(), errors="coerce", downcast="integer")
    except Exception:
        return None

def get_color_by_year(year):
    try:
        y = int(year)
        if y >= 2025:
            return YEAR_COLORS["gte_2025"]
        if y >= 2024:
            return YEAR_COLORS["gte_2024"]
        if y >= 2023:
            return YEAR_COLORS["gte_2023"]
        return YEAR_COLORS["lt_2023"]
    except Exception:
        return "gray"

def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = np.radians(lat1), np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlam = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(phi1) * np.cos(phi2) * np.sin(dlam / 2) ** 2
    return R * 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))

def detect_outliers_iqr(series):
    q1, q3 = series.quantile(0.25), series.quantile(0.75)
    iqr = q3 - q1
    return (series < (q1 - 1.5 * iqr)) | (series > (q3 + 1.5 * iqr))

def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data Tanah")
    return buf.getvalue()

def safe_get(row, col, default="-"):
    val = getattr(row, col, default)
    return default if pd.isna(val) else val

# ─── Sidebar ─────────────────────────────────────────────────────────────────
st.sidebar.markdown(
    '<div class="brand"><div class="logo">SRR</div><div>'
    '<div class="t1">Pangkalan Data Tanah</div>'
    '<div class="t2">KJPP Suwendho Rinaldy dan Rekan</div></div></div>',
    unsafe_allow_html=True,
)
st.sidebar.markdown('<div class="side-label">Sumber Data</div>', unsafe_allow_html=True)
file = st.sidebar.file_uploader("Unggah file Excel data tanah", type=["xlsx"],
                                label_visibility="collapsed")

if file and "last_file" in st.session_state and file != st.session_state["last_file"]:
    for _k in [k for k in st.session_state if k.startswith("adj_df_")]:
        del st.session_state[_k]
    st.session_state.pop("peta_sel", None)
st.session_state["last_file"] = file

if not file:
    st.markdown("""
    <div class="hero">
      <div class="eyebrow">Reviu Penilaian Tanah</div>
      <h1>Pangkalan Data Penilaian Tanah</h1>
      <p>Telaah data pembanding, petakan lokasi terhadap obyek penilaian, dan hitung
         indikasi nilai tanah beserta koefisien variasinya dalam satu tempat.</p>
    </div>
    <div class="feat-grid">
      <div class="feat"><div class="n">01</div><h4>Ringkasan</h4>
        <p>Statistik harga, distribusi, tren per tahun, dan sebaran per kecamatan.</p></div>
      <div class="feat"><div class="n">02</div><h4>Peta Lokasi</h4>
        <p>Obyek dan data pembanding di peta, jarak, foto, serta heatmap harga.</p></div>
      <div class="feat"><div class="n">03</div><h4>Tabel Data</h4>
        <p>Penelusuran data lengkap, deteksi outlier, dan ekspor ke Excel.</p></div>
      <div class="feat"><div class="n">04</div><h4>Analisa Perbandingan</h4>
        <p>Koreksi diskon, waktu, luas, legalitas, lokasi, peruntukan, dan CV.</p></div>
    </div>
    """, unsafe_allow_html=True)
    st.info("Mulai dengan mengunggah file Excel (.xlsx) melalui panel di sebelah kiri.",
            icon=":material/upload_file:")
    with st.expander("Format file yang didukung"):
        st.markdown(
            "| Sheet | Isi | Wajib |\n|---|---|---|\n"
            "| **Data Properti** | Obyek penilaian (field per baris, record per kolom) | Ya* |\n"
            "| **Data Pembanding** | Data pasar pembanding (field per baris, record per kolom) | Ya* |\n"
            "| **Data Bangunan** | Rincian bangunan obyek per kode inspeksi | Opsional |\n"
            "| **Data BTB** | Biaya teknis bangunan per kelas (kolom *Pembulatan*) | Opsional |\n\n"
            "\\* Jika kedua sheet tidak ada, sheet pertama dibaca sebagai tabel datar "
            "dengan kolom `Latitude`, `Longitude`, `Harga_Tanah`, `Luas_Tanah`, dan seterusnya."
        )
    st.stop()

# ─── Helpers untuk loading sheet bertranspose ────────────────────────────────

def parse_indo_number(val):
    """Konversi angka format Indonesia '1.234.567,89' → float 1234567.89"""
    try:
        s = str(val).strip().replace(" ", "").replace("%", "")
        if not s or s in ("nan", "-", ""):
            return None
        if "," in s:
            # Anggap koma = desimal, titik = ribuan
            s = s.replace(".", "").replace(",", ".")
        else:
            # Titik saja: cek apakah itu ribuan (grup 3 angka setelah titik)
            parts = s.split(".")
            if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
                s = s.replace(".", "")
        return float(s)
    except Exception:
        return None

def parse_koordinat(val):
    """Parse '0.640530, 122.907528' → (lat, lon)"""
    try:
        s = str(val).strip()
        parts = s.split(",")
        return float(parts[0].strip()), float(parts[1].strip())
    except Exception:
        return None, None

def _transpose_sheet(file, sheet_name):
    """Baca sheet bertranspose (baris=field, kolom=record) → DataFrame normal."""
    df_raw = pd.read_excel(file, sheet_name=sheet_name, header=0)
    if df_raw.empty:
        return pd.DataFrame()

    # Kolom pertama = nama field; jadikan index setelah dibersihkan duplikatnya
    field_col = df_raw.iloc[:, 0].astype(str).str.strip()

    # Deduplikasi nama field (e.g., "Foto" ke-2 → "Foto_2")
    seen = {}
    unique_fields = []
    for name in field_col:
        if name in seen:
            seen[name] += 1
            unique_fields.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            unique_fields.append(name)

    df_raw = df_raw.iloc[:, 1:]          # Buang kolom field-name
    df_raw.index = unique_fields         # Pasang sebagai index unik
    df = df_raw.T.reset_index(drop=True) # Transpose: record = baris
    df.columns = [str(c).strip() for c in df.columns]
    return df

RENAME_PEMBANDING = {
    "Nomor Data":                       "Nomor",
    "Harga":                            "Harga_Total",
    "Jenis Data":                       "Jenis_Data",
    "Tanggal Perolehan Data":           "Tanggal_Data",
    "Penjual":                          "Kontak",
    "Sumber Data":                      "Sumber_Data",
    "Nomor Telepon Pembanding":         "Telp",
    "Nomor Telepon untuk Konfirmasi":   "Telp_Konfirmasi",
    "Jenis Properti":                   "Jenis_Properti",
    "Alamat":                           "Alamat",
    "Kompleks/Dusun":                   "Kompleks",
    "Desa/Kelurahan":                   "Kelurahan",
    "Kecamatan":                        "Kecamatan",
    "Kabupaten/Kota":                   "Kota",
    "Propinsi":                         "Propinsi",
    "Koordinat":                        "_Koordinat",
    "Luas Tanah":                       "Luas_Tanah",
    "Luas Bangunan":                    "Luas_Bangunan",
    "Kondisi Bangunan":                 "Kondisi_Bangunan",
    "Kelas Bangunan":                   "Kelas_Bangunan",
    "Peruntukan Tata Kota":             "Peruntukan",
    "Bentuk kepemilikan":               "Kepemilikan",
    "Penggunaan Tanah":                 "Penggunaan",
    "Foto Depan Data":                  "Foto",
    "Foto Jalan":                       "Foto_Jalan",
    "Nama Surveyor":                    "Surveyor",
    "Kode Inspeksi":                    "Kode_Inspeksi",
    "Catatan":                          "Catatan",
    "Timestamp":                        "Timestamp",
}

RENAME_PROPERTI = {
    "Timestamp":                        "Timestamp",
    "Kode Inspeksi":                    "Kode_Inspeksi",
    "Nama Surveyor":                    "Surveyor",
    "Tanggal Inspeksi":                 "Tanggal_Inspeksi",
    "Pemberi Tugas":                    "Pemberi_Tugas",
    "Pemilik Properti":                 "Pemilik",
    "Jenis Properti":                   "Jenis_Properti",
    "Alamat":                           "Alamat",
    "Kompleks/Dusun":                   "Kompleks",
    "Desa/Kelurahan":                   "Kelurahan",
    "Kecamatan":                        "Kecamatan",
    "Kabupaten/Kota":                   "Kota",
    "Propinsi":                         "Propinsi",
    "Latitude":                         "Latitude",
    "Longitude":                        "Longitude",
    "Koordinat":                        "_Koordinat",
    "Luas Tanah":                       "Luas_Tanah",
    "Luas Bangunan":                    "Luas_Bangunan",
    "Peruntukan Tata Kota":             "Peruntukan",
    "Bentuk kepemilikan":               "Kepemilikan",
    "Penggunaan Tanah":                 "Penggunaan",
    "Foto Depan Properti":              "Foto",
    "Foto Bagian Dalam":                "Foto_Dalam",
    "Foto Jalan dari Samping Kanan":    "Foto_Samping_Kanan",
    "Foto Jalan dari Samping Kiri":     "Foto_Samping_Kiri",
    "Gambar Situasi dan Plot ATR BPN":  "Gambar_Situasi",
    "Reviewer":                         "Reviewer",
}

RENAME_BANGUNAN = {
    "Kode Inspeksi":    "Kode_Inspeksi",
    "Tanggal Inspeksi": "Tanggal_Inspeksi",
    "Jenis Bangunan":   "Jenis_Bangunan",
    "Luas Bangunan":    "Luas_Bangunan",
    "Nomor Bangunan":   "Nomor_Bangunan",
}

@st.cache_data(show_spinner=False)
def load_bangunan_sheet(uploaded_file):
    try:
        df = _transpose_sheet(uploaded_file, "Data Bangunan")
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    df = df.rename(columns={k: v for k, v in RENAME_BANGUNAN.items() if k in df.columns})
    if "Luas_Bangunan" in df.columns:
        df["Luas_Bangunan"] = df["Luas_Bangunan"].apply(parse_indo_number)
    return df

@st.cache_data(show_spinner=False)
def load_btb_sheet(uploaded_file):
    """Baca sheet Data BTB sebagai tabel flat; return dict {Kelas_Bangunan: Biaya_BTB}."""
    try:
        xl = pd.ExcelFile(uploaded_file)
        if "Data BTB" not in xl.sheet_names:
            return pd.DataFrame()
        df = pd.read_excel(uploaded_file, sheet_name="Data BTB")
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    # Normalkan nama kolom
    df.columns = [str(c).strip() for c in df.columns]
    # Deteksi kolom: "Kelas Bangunan" prioritas, fallback ke "Elemen Bangunan"
    kelas_col = next(
        (c for c in df.columns if "kelas" in c.lower() and "bangunan" in c.lower()), None
    ) or next(
        (c for c in df.columns if "elemen" in c.lower()), None
    )
    biaya_col = next(
        (c for c in df.columns if "pembulatan" in c.lower()), None
    )
    rename_btb = {}
    if kelas_col:
        rename_btb[kelas_col] = "Kelas_Bangunan"
    if biaya_col:
        rename_btb[biaya_col] = "Biaya_BTB"
    df = df.rename(columns=rename_btb)
    if "Biaya_BTB" in df.columns:
        df["Biaya_BTB"] = df["Biaya_BTB"].apply(parse_indo_number)
    return df

def load_pembanding_sheet(file):
    try:
        df = _transpose_sheet(file, "Data Pembanding")
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()

    df = df.rename(columns={k: v for k, v in RENAME_PEMBANDING.items() if k in df.columns})

    # Koordinat → Latitude, Longitude
    if "_Koordinat" in df.columns:
        coords = df["_Koordinat"].apply(
            lambda x: pd.Series(parse_koordinat(x), index=["Latitude", "Longitude"])
        )
        df["Latitude"]  = coords["Latitude"]
        df["Longitude"] = coords["Longitude"]

    # Numerik
    for col in ["Harga_Total", "Luas_Tanah", "Luas_Bangunan", "Kondisi_Bangunan"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_indo_number)

    # Harga per m² — sementara pakai rumus sederhana; akan dikoreksi BTB setelah load
    if "Harga_Total" in df.columns and "Luas_Tanah" in df.columns:
        df["Harga_Tanah"] = (df["Harga_Total"] / df["Luas_Tanah"]).round(0)

    # Tahun dari Tanggal_Data atau Timestamp
    for dcol in ["Tanggal_Data", "Timestamp"]:
        if dcol in df.columns:
            df["Tahun"] = pd.to_datetime(
                df[dcol].astype(str), errors="coerce", dayfirst=True
            ).dt.year
            break

    # Pastikan Nomor ada
    if "Nomor" not in df.columns:
        df["Nomor"] = range(1, len(df) + 1)

    df["_sumber"] = "Data Pembanding"
    return df

def load_properti_sheet(file):
    try:
        df = _transpose_sheet(file, "Data Properti")
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()

    df = df.rename(columns={k: v for k, v in RENAME_PROPERTI.items() if k in df.columns})

    # Koordinat fallback
    if ("Latitude" not in df.columns or df["Latitude"].isna().all()) and "_Koordinat" in df.columns:
        coords = df["_Koordinat"].apply(
            lambda x: pd.Series(parse_koordinat(x), index=["Latitude", "Longitude"])
        )
        df["Latitude"]  = coords["Latitude"]
        df["Longitude"] = coords["Longitude"]

    for col in ["Latitude", "Longitude"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    for col in ["Luas_Tanah", "Luas_Bangunan"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_indo_number)

    # Tahun dari Tanggal_Inspeksi atau Timestamp
    for dcol in ["Tanggal_Inspeksi", "Timestamp"]:
        if dcol in df.columns:
            df["Tahun"] = pd.to_datetime(
                df[dcol].astype(str), errors="coerce", dayfirst=True
            ).dt.year
            break

    df["Nomor"]      = "Obyek Penilaian"
    df["Harga_Tanah"] = np.nan
    df["Harga_Total"] = np.nan
    df["_sumber"]    = "Data Properti"
    return df

@st.cache_data(show_spinner="Memuat data...")
def load_data(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    sheets = xl.sheet_names

    has_pembanding = "Data Pembanding" in sheets
    has_properti   = "Data Properti"   in sheets

    if has_pembanding or has_properti:
        frames = []
        if has_properti:
            dp = load_properti_sheet(uploaded_file)
            if not dp.empty:
                frames.append(dp)
        if has_pembanding:
            db = load_pembanding_sheet(uploaded_file)
            if not db.empty:
                frames.append(db)
        if frames:
            df = pd.concat(frames, ignore_index=True)
            if "Latitude"  in df.columns:
                df["Latitude"]  = pd.to_numeric(df["Latitude"],  errors="coerce")
            if "Longitude" in df.columns:
                df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
            df["_format"]   = "multi-sheet"
            return df

    # Fallback: format lama (flat sheet pertama)
    df = pd.read_excel(uploaded_file)
    df["Latitude"]  = pd.to_numeric(df["Latitude"],  errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df["_format"]   = "flat"
    return df

df = load_data(file)
df_bangunan = load_bangunan_sheet(file)
df_btb      = load_btb_sheet(file)
df["Tahun_Bersih"] = df["Tahun"].apply(bersihkan_tahun) if "Tahun" in df.columns else pd.Series(dtype=float)

# ── Koreksi Harga_Tanah dengan ekstraksi nilai bangunan (BTB) ──────────────
# Rumus: (Harga_Total − Luas_Bangunan × (Kondisi/100) × Biaya_BTB) / Luas_Tanah
_btb_msg = ""
if df_btb.empty:
    _btb_msg = "⚠️ Sheet 'Data BTB' tidak ditemukan — harga/m² pakai Harga_Total / Luas_Tanah"
elif "Kelas_Bangunan" not in df_btb.columns:
    _btb_msg = f"⚠️ Kolom kelas bangunan tidak dikenali di BTB (kolom: {', '.join(df_btb.columns.tolist()[:6])})"
elif "Biaya_BTB" not in df_btb.columns:
    _btb_msg = f"⚠️ Kolom 'Pembulatan' tidak ditemukan di BTB (kolom: {', '.join(df_btb.columns.tolist()[:6])})"
else:
    # Normalisasi key ke lowercase+strip agar pencocokan tidak case-sensitive
    btb_map = (
        df_btb
        .dropna(subset=["Kelas_Bangunan", "Biaya_BTB"])
        .assign(Kelas_Bangunan=lambda x: x["Kelas_Bangunan"].astype(str).str.strip().str.lower())
        .set_index("Kelas_Bangunan")["Biaya_BTB"]
        .to_dict()
    )
    has_sumber = "_sumber" in df.columns
    mask_comp  = (df["_sumber"] == "Data Pembanding") if has_sumber else pd.Series(True, index=df.index)

    if mask_comp.any() and "Kelas_Bangunan" in df.columns:
        df_c = df[mask_comp].copy()
        df_c["_Kelas_norm"] = df_c["Kelas_Bangunan"].astype(str).str.strip().str.lower()
        df_c["_Biaya_BTB"]  = df_c["_Kelas_norm"].map(btb_map)

        need     = ["Harga_Total", "Luas_Bangunan", "Kondisi_Bangunan", "Luas_Tanah", "_Biaya_BTB"]
        miss_col = [c for c in need[:-1] if c not in df_c.columns]
        if miss_col:
            _btb_msg = f"⚠️ Kolom tidak ada di data: {', '.join(miss_col)}"
        else:
            has_all   = df_c[need].notna().all(axis=1) & (df_c["Luas_Tanah"] > 0)
            n_nomatch = int((df_c["_Biaya_BTB"].isna() &
                             df_c[["Harga_Total","Luas_Bangunan","Kondisi_Bangunan","Luas_Tanah"]].notna().all(axis=1)).sum())
            if has_all.any():
                sub       = df_c[has_all]
                nilai_bgn = sub["Luas_Bangunan"] * (sub["Kondisi_Bangunan"] / 100) * sub["_Biaya_BTB"]
                harga_baru = ((sub["Harga_Total"] - nilai_bgn) / sub["Luas_Tanah"]).round(0)
                df.loc[harga_baru.index, "Harga_Tanah"] = harga_baru
                _btb_msg = f"✅ BTB koreksi diterapkan pada {int(has_all.sum())} data pembanding"
                if n_nomatch:
                    unmatched = df_c.loc[df_c["_Biaya_BTB"].isna(), "Kelas_Bangunan"].dropna().unique().tolist()
                    _btb_msg += f" | ⚠️ {n_nomatch} kelas tidak cocok: {unmatched}"
            else:
                unmatched = df_c.loc[df_c["_Biaya_BTB"].isna(), "Kelas_Bangunan"].dropna().unique().tolist()
                btb_keys  = list(btb_map.keys())
                _btb_msg  = (f"⚠️ BTB dimuat ({len(btb_map)} kelas) tapi tidak ada yang cocok.\n"
                             f"Kelas di data: {unmatched[:4]}\n"
                             f"Kelas di BTB : {btb_keys[:4]}")
    else:
        _btb_msg = "ℹ️ Tidak ada data pembanding atau kolom Kelas_Bangunan tidak ada"

# ─── Status data di sidebar ──────────────────────────────────────────────────
n_prop = int((df["_sumber"] == "Data Properti").sum())   if "_sumber" in df.columns else 0
n_pemb = int((df["_sumber"] == "Data Pembanding").sum()) if "_sumber" in df.columns else len(df)

def _status_line(color, text):
    return (f'<div class="status-line"><span class="dot" style="background:{color}"></span>'
            f'<span>{text}</span></div>')

_status = [_status_line(TEAL, f"<b>{n_prop}</b> obyek penilaian · <b>{n_pemb}</b> data pembanding")]
_btb_ok = _btb_msg.startswith("✅") and "⚠️" not in _btb_msg
if _btb_msg:
    _btb_short = (_btb_msg.split("|")[0].strip().lstrip("✅").strip() if _btb_msg.startswith("✅")
                  else "Koreksi BTB tidak diterapkan")
    _status.append(_status_line(TEAL if _btb_ok else AMBER, _btb_short))
st.sidebar.markdown("".join(_status), unsafe_allow_html=True)
if _btb_msg and not _btb_ok:
    with st.sidebar.expander("Detail koreksi BTB"):
        st.caption(_btb_msg.replace("✅", "").replace("⚠️", "").replace("ℹ️", "").strip())

# ─── Filter controls ──────────────────────────────────────────────────────────
st.sidebar.markdown('<div class="side-label">Filter Data</div>', unsafe_allow_html=True)
with st.sidebar.form("filter_form", border=False):
    city_input = st.text_input("Kabupaten/Kota", placeholder="Ketik sebagian nama…")

    available_years = sorted([int(y) for y in df["Tahun_Bersih"].dropna().unique()], reverse=True)
    selected_year = st.selectbox("Tahun data", ["Semua Tahun"] + [str(y) for y in available_years])

    kec_options = ["Semua Kecamatan"] + sorted(df["Kecamatan"].dropna().astype(str).unique().tolist())
    selected_kecamatan = st.selectbox("Kecamatan", kec_options)

    harga_col = df["Harga_Tanah"].dropna()
    price_range = None
    if not harga_col.empty and harga_col.min() < harga_col.max():
        h_min, h_max = float(harga_col.min()), float(harga_col.max())
        price_range = st.slider("Harga tanah (Rp/m²)", h_min, h_max, (h_min, h_max), format="%.0f")

    luas_col = df["Luas_Tanah"].dropna()
    luas_range = None
    if not luas_col.empty and luas_col.min() < luas_col.max():
        l_min, l_max = float(luas_col.min()), float(luas_col.max())
        luas_range = st.slider("Luas tanah (m²)", l_min, l_max, (l_min, l_max), format="%.0f")

    st.form_submit_button("Terapkan Filter", type="primary", use_container_width=True,
                          icon=":material/filter_alt:")
st.sidebar.caption("Obyek penilaian selalu ditampilkan, apa pun filternya.")

# ─── Apply filters ────────────────────────────────────────────────────────────
filtered = df.copy()

# Identifikasi baris Obyek Penilaian — selalu dipertahankan di semua filter
_is_obyek = filtered["Nomor"].astype(str).str.strip().str.lower().str.contains("obyek", na=False)

if city_input.strip():
    _city_ok = filtered["Kota"].astype(str).str.strip().str.lower().str.contains(
        city_input.strip().lower(), na=False
    )
    filtered = filtered[_city_ok | _is_obyek]
    _is_obyek = filtered["Nomor"].astype(str).str.strip().str.lower().str.contains("obyek", na=False)

if selected_year != "Semua Tahun":
    _year_ok = filtered["Tahun_Bersih"] == int(selected_year)
    filtered = filtered[_year_ok | _is_obyek]
    _is_obyek = filtered["Nomor"].astype(str).str.strip().str.lower().str.contains("obyek", na=False)

if selected_kecamatan != "Semua Kecamatan":
    _kec_ok = filtered["Kecamatan"].astype(str) == selected_kecamatan
    filtered = filtered[_kec_ok | _is_obyek]
    _is_obyek = filtered["Nomor"].astype(str).str.strip().str.lower().str.contains("obyek", na=False)

if price_range:
    # Obyek Penilaian sering tidak punya Harga_Tanah — jangan dibuang
    _price_ok = (
        filtered["Harga_Tanah"].isna() |
        ((filtered["Harga_Tanah"] >= price_range[0]) &
         (filtered["Harga_Tanah"] <= price_range[1]))
    )
    filtered = filtered[_price_ok | _is_obyek]
    _is_obyek = filtered["Nomor"].astype(str).str.strip().str.lower().str.contains("obyek", na=False)

if luas_range:
    _luas_ok = (
        filtered["Luas_Tanah"].isna() |
        ((filtered["Luas_Tanah"] >= luas_range[0]) &
         (filtered["Luas_Tanah"] <= luas_range[1]))
    )
    filtered = filtered[_luas_ok | _is_obyek]

# Outlier flag
valid_harga = filtered["Harga_Tanah"].dropna()
if len(valid_harga) >= 4:
    filtered = filtered.copy()
    filtered["_outlier"] = False
    outlier_mask = detect_outliers_iqr(filtered.loc[filtered["Harga_Tanah"].notna(), "Harga_Tanah"])
    filtered.loc[outlier_mask.index, "_outlier"] = outlier_mask.values
else:
    filtered = filtered.copy()
    filtered["_outlier"] = False

city_label = city_input.strip() or "Semua Kota"

# ─── Header ───────────────────────────────────────────────────────────────────
_n_comp_f = int((~filtered["Nomor"].astype(str).str.lower().str.contains("obyek", na=False)).sum())
_chips = [
    f"<span class='chip'><b>{_n_comp_f}</b> data pembanding</span>",
    f"<span class='chip'>Kota: <b>{city_label}</b></span>",
    f"<span class='chip'>Tahun: <b>{selected_year.replace('Semua Tahun', 'Semua')}</b></span>",
]
if selected_kecamatan != "Semua Kecamatan":
    _chips.append(f"<span class='chip'>Kec.: <b>{selected_kecamatan}</b></span>")
st.markdown(
    f'<div class="app-header"><div><h1>Reviu Penilaian Tanah</h1>'
    f'<div class="sub">{file.name}</div></div>'
    f'<div class="chips">{"".join(_chips)}</div></div>',
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("Tidak ada data yang sesuai dengan filter yang dipilih.", icon=":material/search_off:")
    st.stop()

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_dashboard, tab_peta, tab_tabel, tab_analisa = st.tabs([
    ":material/insights: Ringkasan",
    ":material/map: Peta Lokasi",
    ":material/table_rows: Tabel Data",
    ":material/calculate: Analisa Perbandingan",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — RINGKASAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_dashboard:
    prices = filtered["Harga_Tanah"].dropna()
    outlier_count = int(filtered["_outlier"].sum())

    def _rp_or_na(v):
        return rp(v) if not prices.empty else "—"
    _cv_mkt = prices.std() / prices.mean() * 100 if len(prices) > 1 and prices.mean() else None
    kpi_cards([
        ("Rata-rata harga", _rp_or_na(prices.mean()), f"{len(prices)} data berharga", True),
        ("Median",          _rp_or_na(prices.median()), None, False),
        ("Minimum",         _rp_or_na(prices.min()), None, False),
        ("Maksimum",        _rp_or_na(prices.max()), None, False),
        ("Sebaran (CV)",    f"{_cv_mkt:.1f}%" if _cv_mkt is not None else "—",
                            f"{outlier_count} outlier (IQR 1,5×)" if outlier_count else "Tanpa outlier", False),
    ])

    if outlier_count:
        st.caption(f":material/warning: {outlier_count} data terdeteksi sebagai outlier — "
                   "rinciannya tersedia di tab **Tabel Data**.")

    col_l, col_r = st.columns(2, gap="medium")

    with col_l:
        # Distribusi harga
        with st.container(border=True):
            if not prices.empty:
                fig_hist = px.histogram(
                    filtered[filtered["Harga_Tanah"].notna()],
                    x="Harga_Tanah", nbins=20,
                    labels={"Harga_Tanah": "Harga (Rp/m²)"},
                    color_discrete_sequence=[NAVY],
                )
                fig_hist.update_traces(marker_line_color="white", marker_line_width=1)
                fig_hist.update_layout(showlegend=False, bargap=0.05, yaxis_title="Jumlah data")
                st.plotly_chart(style_fig(fig_hist, "Distribusi harga tanah"), use_container_width=True, config=PLOT_CFG)
            else:
                st.caption("Belum ada data harga.")

        # Scatter harga vs luas
        scatter_df = filtered[filtered["Harga_Tanah"].notna() & filtered["Luas_Tanah"].notna()].copy()
        if not scatter_df.empty:
            scatter_df["Tahun_Label"] = scatter_df["Tahun_Bersih"].astype("Int64").astype(str)
            with st.container(border=True):
                fig_scatter = px.scatter(
                    scatter_df.sort_values("Tahun_Label"), x="Luas_Tanah", y="Harga_Tanah",
                    color="Tahun_Label",
                    hover_data=["Nomor", "Alamat", "Kecamatan"],
                    labels={"Luas_Tanah": "Luas tanah (m²)", "Harga_Tanah": "Harga (Rp/m²)",
                            "Tahun_Label": "Tahun"},
                    trendline="ols", trendline_scope="overall",
                    trendline_color_override=SLATE,
                    color_discrete_sequence=PLOT_SEQ,
                )
                fig_scatter.update_traces(marker=dict(size=9, line=dict(width=1, color="white")),
                                          selector=dict(mode="markers"))
                fig_scatter.update_traces(name="Tren (OLS)", selector=dict(mode="lines"))
                st.plotly_chart(style_fig(fig_scatter, "Harga terhadap luas tanah"),
                                use_container_width=True, config=PLOT_CFG)

    with col_r:
        # Tren harga per tahun
        by_year = (
            filtered.groupby("Tahun_Bersih")["Harga_Tanah"]
            .agg(rata_rata="mean", minimum="min", maksimum="max", jumlah="count")
            .reset_index()
            .rename(columns={"Tahun_Bersih": "Tahun"})
            .dropna(subset=["Tahun"])
            .sort_values("Tahun")
        )
        if not by_year.empty:
            with st.container(border=True):
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Scatter(
                    x=by_year["Tahun"], y=by_year["maksimum"], mode="lines", name="Maksimum",
                    line=dict(width=0), showlegend=False, hoverinfo="skip",
                ))
                fig_trend.add_trace(go.Scatter(
                    x=by_year["Tahun"], y=by_year["minimum"], mode="lines", name="Rentang min–maks",
                    line=dict(width=0), fill="tonexty", fillcolor="rgba(31,78,121,0.10)",
                ))
                fig_trend.add_trace(go.Scatter(
                    x=by_year["Tahun"], y=by_year["rata_rata"],
                    mode="lines+markers", name="Rata-rata",
                    line=dict(color=NAVY, width=2.5), marker=dict(size=7),
                    customdata=by_year[["minimum", "maksimum", "jumlah"]],
                    hovertemplate="Rata-rata: Rp %{y:,.0f}<br>Min: Rp %{customdata[0]:,.0f}"
                                  "<br>Maks: Rp %{customdata[1]:,.0f}<br>%{customdata[2]} data<extra></extra>",
                ))
                fig_trend.update_layout(xaxis_title="Tahun", yaxis_title="Harga (Rp/m²)",
                                        hovermode="x unified")
                fig_trend.update_xaxes(dtick=1, tickformat="d")
                st.plotly_chart(style_fig(fig_trend, "Tren harga per tahun"), use_container_width=True, config=PLOT_CFG)

        # Rata-rata harga per kecamatan
        by_kec = (
            filtered.groupby("Kecamatan")["Harga_Tanah"]
            .agg(["mean", "count"])
            .dropna()
            .sort_values("mean", ascending=True)
            .reset_index()
        )
        if not by_kec.empty:
            with st.container(border=True):
                fig_kec = go.Figure(go.Bar(
                    x=by_kec["mean"], y=by_kec["Kecamatan"], orientation="h",
                    marker_color=NAVY, customdata=by_kec["count"],
                    hovertemplate="%{y}: Rp %{x:,.0f}/m² (%{customdata} data)<extra></extra>",
                ))
                fig_kec.update_layout(xaxis_title="Rata-rata harga (Rp/m²)", yaxis_title=None,
                                      bargap=0.35)
                style_fig(fig_kec, "Rata-rata harga per kecamatan",
                          height=max(340, 60 + 28 * len(by_kec)))
                fig_kec.update_xaxes(showgrid=True, gridcolor="#EEF2F6", separatethousands=True)
                fig_kec.update_yaxes(showgrid=False)
                st.plotly_chart(fig_kec, use_container_width=True, config=PLOT_CFG)

    # Ringkasan statistik
    num_cols = [c for c in ["Harga_Tanah", "Luas_Tanah", "Luas_Bangunan"] if c in filtered.columns]
    if num_cols:
        with st.expander("Ringkasan statistik", icon=":material/functions:"):
            stats = filtered[num_cols].describe()
            stats.index = ["Jumlah", "Rata-rata", "Std. deviasi", "Minimum",
                           "Kuartil 1", "Median", "Kuartil 3", "Maksimum"]
            stats.columns = [COL_LABELS.get(c, c) for c in stats.columns]
            st.dataframe(
                stats, use_container_width=True,
                column_config={c: st.column_config.NumberColumn(format="localized") for c in stats.columns},
            )

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PETA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_peta:
    map_df = filtered[filtered["Latitude"].notna() & filtered["Longitude"].notna()].copy()
    map_df = map_df[
        map_df["Latitude"].between(-90, 90) &
        map_df["Longitude"].between(-180, 180)
    ]

    subj_mask   = map_df["Nomor"].astype(str).str.lower().str.contains("obyek", na=False)
    subj_df     = map_df[subj_mask]
    comp_map_df = map_df[~subj_mask]

    # ── Toolbar peta ──────────────────────────────────────────────────────────
    tb1, tb2, tb3, tb4 = st.columns([3, 1.2, 1.2, 1.2], vertical_alignment="center")
    tb1.caption(f"**{len(subj_df)}** obyek penilaian · **{len(comp_map_df)}** data pembanding "
                "berkoordinat — klik penanda untuk melihat detail.")
    show_lines   = tb2.toggle("Garis jarak", value=True)
    show_labels  = tb3.toggle("Label harga", value=True)
    show_heatmap = tb4.toggle("Heatmap harga", value=False)

    # Pusatkan peta ke obyek penilaian jika ada, atau rata-rata semua
    if not subj_df.empty:
        lat0, lon0 = float(subj_df.iloc[0]["Latitude"]), float(subj_df.iloc[0]["Longitude"])
    elif not map_df.empty:
        lat0, lon0 = map_df["Latitude"].mean(), map_df["Longitude"].mean()
    else:
        lat0, lon0 = -2.548926, 118.0148634

    m = folium.Map(
        location=[lat0, lon0],
        zoom_start=14 if len(map_df) <= 20 else 12,
        min_zoom=5, max_zoom=18,
        prefer_canvas=True, control_scale=True,
        tiles=None,
    )
    folium.TileLayer(
        tiles="https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png",
        name="Peta jalan", attr="©OpenStreetMap ©CartoDB",
    ).add_to(m)
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        name="Citra satelit", attr="Tiles © Esri",
    ).add_to(m)

    if show_heatmap and not comp_map_df.empty:
        heat = [
            [r.Latitude, r.Longitude, float(r.Harga_Tanah)]
            for r in comp_map_df.itertuples()
            if pd.notna(getattr(r, "Harga_Tanah", None))
        ]
        if heat:
            HeatMap(heat, name="Heatmap harga", radius=30, blur=20, min_opacity=0.4).add_to(m)

    # ── Garis jarak dari Obyek Penilaian ke setiap Data Pembanding ──────────
    if show_lines and not subj_df.empty and not comp_map_df.empty:
        s = subj_df.iloc[0]
        lines_fg = folium.FeatureGroup(name="Garis jarak", show=True).add_to(m)

        for r in comp_map_df.itertuples():
            nomor_r = str(safe_get(r, "Nomor")).strip()
            dist_km = haversine_km(s.Latitude, s.Longitude, r.Latitude, r.Longitude)
            dist_label = f"{dist_km:.2f} km" if dist_km >= 1 else f"{dist_km*1000:.0f} m"

            folium.PolyLine(
                locations=[[s.Latitude, s.Longitude], [r.Latitude, r.Longitude]],
                color=NAVY_DK, weight=1.5, dash_array="5 5", opacity=0.7,
                tooltip=f"Obyek → Data {nomor_r}: {dist_label}",
            ).add_to(lines_fg)

            # Label jarak di tengah garis
            folium.Marker(
                [(s.Latitude + r.Latitude) / 2, (s.Longitude + r.Longitude) / 2],
                icon=folium.DivIcon(
                    html=f"""<div style="font:600 10px/1.2 'Segoe UI',sans-serif;color:{NAVY_DK};
                                background:rgba(255,255,255,.92);padding:2px 6px;border-radius:9px;
                                border:1px solid {LINE};white-space:nowrap;pointer-events:none;
                                display:inline-block">{dist_label}</div>""",
                    icon_size=(70, 18), icon_anchor=(35, 9),
                ),
            ).add_to(lines_fg)

    marker_layer = folium.FeatureGroup(name="Data pembanding").add_to(m)
    subj_layer   = folium.FeatureGroup(name="Obyek penilaian", show=True).add_to(m)

    for r in map_df.itertuples():
        nomor     = str(safe_get(r, "Nomor")).strip()
        tahun     = getattr(r, "Tahun_Bersih", 0) or 0
        harga_fmt = format_currency(getattr(r, "Harga_Tanah", 0))

        if "obyek" in nomor.lower():
            folium.Marker(
                location=[r.Latitude, r.Longitude],
                tooltip="Obyek Penilaian",
                icon=folium.DivIcon(
                    html=f"""<div style="width:34px;height:34px;border-radius:50% 50% 50% 0;
                                transform:rotate(-45deg);background:{NAVY_DK};border:3px solid #fff;
                                box-shadow:0 2px 8px rgba(0,0,0,.45);display:flex;
                                align-items:center;justify-content:center">
                               <span style="transform:rotate(45deg);color:#fff;
                                 font:700 11px 'Segoe UI',sans-serif">OP</span></div>""",
                    icon_size=(34, 34), icon_anchor=(17, 40),
                ),
            ).add_to(subj_layer)
        else:
            warna = get_color_by_year(tahun)
            folium.Marker(
                location=[r.Latitude, r.Longitude],
                tooltip=f"Data {nomor} · {harga_fmt}/m²",
                icon=folium.DivIcon(
                    html=f"""<div style="width:26px;height:26px;border-radius:50%;background:{warna};
                                border:2.5px solid #fff;box-shadow:0 1px 5px rgba(0,0,0,.4);
                                display:flex;align-items:center;justify-content:center;
                                color:#fff;font:700 10px 'Segoe UI',sans-serif">{nomor}</div>""",
                    icon_size=(26, 26), icon_anchor=(13, 13),
                ),
            ).add_to(marker_layer)

            if show_labels:
                folium.Marker(
                    [r.Latitude, r.Longitude],
                    icon=folium.DivIcon(
                        html=f"""<div style="font:600 10px/1.25 'Segoe UI',sans-serif;color:#1E293B;
                                    background:rgba(255,255,255,.88);padding:2px 6px;border-radius:4px;
                                    border-left:3px solid {warna};white-space:nowrap;
                                    pointer-events:none;display:inline-block;
                                    box-shadow:0 1px 3px rgba(0,0,0,.15)">{harga_fmt}/m²</div>""",
                        icon_size=(120, 18), icon_anchor=(-16, 9),
                    ),
                ).add_to(marker_layer)

    folium.LayerControl(collapsed=True).add_to(m)

    def _lg(color, text, shape="50%"):
        return (f'<div style="display:flex;align-items:center;gap:7px;margin:3px 0">'
                f'<span style="width:11px;height:11px;border-radius:{shape};background:{color};'
                f'display:inline-block;border:1.5px solid #fff;box-shadow:0 0 0 1px {LINE}"></span>{text}</div>')
    legend = (
        '<div style="position:fixed;bottom:24px;left:12px;z-index:1000;background:rgba(255,255,255,.96);'
        f'padding:9px 12px;border-radius:8px;border:1px solid {LINE};'
        "font:12px 'Segoe UI',sans-serif;color:#334155;box-shadow:0 2px 8px rgba(0,0,0,.12)\">"
        '<div style="font-weight:700;margin-bottom:4px;color:#0F2A44">Tahun data</div>'
        + _lg(YEAR_COLORS["gte_2025"], "2025 ke atas")
        + _lg(YEAR_COLORS["gte_2024"], "2024")
        + _lg(YEAR_COLORS["gte_2023"], "2023")
        + _lg(YEAR_COLORS["lt_2023"], "Sebelum 2023")
        + _lg(NAVY_DK, "Obyek penilaian", "50% 50% 50% 0")
        + '</div>'
    )
    m.get_root().html.add_child(folium.Element(legend))

    # ── Split layout: peta kiri | panel detail kanan ────────────────────────
    col_map, col_detail = st.columns([7, 3], gap="medium")

    with col_map:
        result = st_folium(
            m, width="100%", height=640,
            returned_objects=["last_object_clicked"],
            key="folium_peta",
        )
        if result and result.get("last_object_clicked"):
            c = result["last_object_clicked"]
            lat_c, lng_c = c.get("lat"), c.get("lng")
            if lat_c is not None and lng_c is not None and not map_df.empty:
                dists = (
                    (map_df["Latitude"]  - lat_c) ** 2 +
                    (map_df["Longitude"] - lng_c) ** 2
                )
                if dists.min() < 1e-5:
                    st.session_state["peta_sel"] = int(dists.idxmin())

    # ── Panel detail ─────────────────────────────────────────────────────────
    with col_detail:
        sel_idx = st.session_state.get("peta_sel")
        if sel_idx is None or sel_idx not in map_df.index:
            components.html(
                f'<div style="height:632px;display:flex;align-items:center;justify-content:center;'
                f'border:1px dashed #CBD5E1;border-radius:10px;background:#F8FAFC;'
                f"font-family:'Source Sans Pro','Segoe UI',sans-serif;color:{SLATE}\">"
                '<div style="text-align:center;padding:0 24px">'
                f'<div style="font-weight:600;color:{NAVY_DK};font-size:14px">Belum ada yang dipilih</div>'
                '<div style="font-size:12.5px;margin-top:6px;line-height:1.5">'
                'Klik penanda obyek atau data pembanding di peta untuk melihat foto dan rinciannya.</div>'
                '</div></div>',
                height=640)
        else:
            row      = map_df.loc[sel_idx]
            nomor_d  = str(safe_get(row, "Nomor")).strip()
            is_s     = "obyek" in nomor_d.lower()
            tahun_d  = row.get("Tahun_Bersih") if hasattr(row, "get") else None
            harga_d  = format_currency(row.get("Harga_Tanah"))
            foto_url = str(safe_get(row, "Foto", ""))

            # ── Kumpulkan semua foto ──────────────────────────────────────
            all_fotos = []
            for u_raw, lb in (
                [(foto_url, "Depan"),
                 (str(safe_get(row, "Foto_Jalan", "")), "Jalan")]
                + ([(str(safe_get(row, c, "")), lb2)
                    for c, lb2 in [("Foto_Dalam","Dalam"),
                                   ("Foto_Samping_Kanan","Kanan"),
                                   ("Foto_Samping_Kiri","Kiri"),
                                   ("Gambar_Situasi","Situasi")]]
                   if is_s else [])
            ):
                if u_raw not in ("", "#", "-", "nan", "None"):
                    lh, th = gdrive_thumbnail(u_raw, width=600)
                    if lh:
                        all_fotos.append((lh, th, u_raw, lb))

            # ── Helper builders ───────────────────────────────────────────
            def _v(x):
                s = str(x) if x is not None else ""
                return s if s not in ("-", "", "None", "nan") else "—"
            def sh(title):
                return (f'<div style="color:{SLATE};font-size:10.5px;font-weight:700;'
                        f'text-transform:uppercase;letter-spacing:.06em;margin:14px 0 4px;'
                        f'padding-bottom:4px;border-bottom:1px solid {LINE}">{title}</div>')
            def r1(lbl, val):
                return (f'<div style="padding:4px 0">'
                        f'<div style="color:{SLATE};font-size:11px">{lbl}</div>'
                        f'<div style="color:#1E293B;font-size:13px;font-weight:500">{_v(val)}</div></div>')
            def r2(l1, v1, l2, v2):
                return (f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:0 10px">'
                        f'{r1(l1, v1)}{r1(l2, v2)}</div>')
            def _luas(col):
                v = row.get(col)
                if v is None or (isinstance(v, float) and pd.isna(v)): return "—"
                try: return f"{float(v):,.0f} m²".replace(",", ".")
                except: return str(v)

            # ── Build HTML ────────────────────────────────────────────────
            P = []
            P.append('<div style="height:632px;overflow-y:auto;padding:0 6px 8px 0;'
                     "font-family:'Source Sans Pro','Segoe UI',sans-serif;color:#1E293B;"
                     'scrollbar-width:thin;scrollbar-color:#CBD5E1 transparent">')

            # Header
            if is_s:
                P.append(f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">'
                         f'<span style="background:{NAVY_DK};color:#fff;border-radius:6px;padding:3px 8px;'
                         f'font-size:11px;font-weight:700">OP</span>'
                         f'<b style="font-size:16px;color:{NAVY_DK}">Obyek Penilaian</b></div>')
            else:
                jenis_d = safe_get(row, "Jenis_Data")
                is_pen  = "penawaran" in str(jenis_d).lower()
                b_bg, b_fg = ("#FDF3E1", "#9A6412") if is_pen else ("#E6EEF6", NAVY)
                thn_lbl = f"{int(tahun_d)}" if tahun_d and not pd.isna(tahun_d) else ""
                P.append(f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">'
                         f'<b style="font-size:16px;color:{NAVY_DK}">Data {nomor_d}</b>'
                         f'<span style="background:{b_bg};color:{b_fg};border-radius:999px;'
                         f'padding:2px 9px;font-size:11px;font-weight:600">{jenis_d}</span>'
                         f'<span style="color:{SLATE};font-size:12px;margin-left:auto">{thn_lbl}</span></div>')

            # Carousel foto — menggunakan JavaScript (berjalan di components.html iframe)
            if all_fotos:
                n_foto  = len(all_fotos)
                slides  = ",".join(
                    f'{{s:"{lh}",f:"{th}",h:"{hr}",l:"{lb}"}}'
                    for lh, th, hr, lb in all_fotos)
                btn_css = ("background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.35);"
                           "color:#fff;font-size:17px;line-height:1;width:28px;height:28px;"
                           "border-radius:6px;cursor:pointer;flex-shrink:0")
                P.append(f'''
<div style="position:relative;margin-bottom:8px;border-radius:8px;overflow:hidden;background:#0F172A">
  <img id="simg" src="{all_fotos[0][0]}"
       style="width:100%;height:auto;display:block"
       referrerpolicy="no-referrer"
       onerror="this.src='{all_fotos[0][1]}';this.onerror=null;">
  <div style="position:absolute;bottom:0;left:0;right:0;
              background:linear-gradient(transparent,rgba(0,0,0,.65));
              padding:6px 8px;display:flex;align-items:center;gap:6px">
    <button id="btnP" style="{btn_css}">&#8249;</button>
    <span id="slbl" style="color:#fff;font-size:11px;font-weight:600;
                           flex:1;text-align:center">Depan 1/{n_foto}</span>
    <button id="btnN" style="{btn_css}">&#8250;</button>
    <a id="slnk" href="{all_fotos[0][2]}" target="_blank"
       style="color:rgba(255,255,255,.85);font-size:11px;
              text-decoration:none;white-space:nowrap;flex-shrink:0">Buka ↗</a>
  </div>
</div>
<script>
(function(){{
  var D=[{slides}], i=0;
  function go(n){{
    i=(n+D.length)%D.length;
    var el=document.getElementById('simg');
    el.src=D[i].s;
    el.onerror=function(){{el.src=D[i].f;el.onerror=null;}};
    document.getElementById('slbl').textContent=D[i].l+' '+(i+1)+'/{n_foto}';
    document.getElementById('slnk').href=D[i].h;
  }}
  document.getElementById('btnP').addEventListener('click',function(){{go(i-1);}});
  document.getElementById('btnN').addEventListener('click',function(){{go(i+1);}});
}})();
</script>''')

            # Link Street View
            sv = generate_streetview_url(row.get("Latitude"), row.get("Longitude"))
            foto_hint = f'{len(all_fotos)} foto' if all_fotos else 'Tidak ada foto'
            P.append(f'<div style="display:flex;justify-content:space-between;align-items:center;'
                     f'margin:2px 0 6px;font-size:12px">'
                     f'<a href="{sv}" target="_blank" style="color:{NAVY};text-decoration:none;'
                     f'font-weight:600">Buka Street View ↗</a>'
                     f'<span style="color:{SLATE}">{foto_hint}</span></div>')

            # Harga (Data Pembanding)
            if not is_s:
                ht     = row.get("Harga_Total")
                ht_str = format_currency(ht) if ht and not pd.isna(ht) else "—"
                box = (f'background:#F8FAFC;border:1px solid {LINE};border-radius:8px;padding:8px 10px')
                P.append(f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:6px 0">'
                         f'<div style="{box}"><div style="font-size:10.5px;color:{SLATE};'
                         f'text-transform:uppercase;letter-spacing:.05em">Harga total</div>'
                         f'<div style="font-size:13.5px;font-weight:700;color:#1E293B">{ht_str}</div></div>'
                         f'<div style="{box};background:{NAVY};border-color:{NAVY}">'
                         f'<div style="font-size:10.5px;color:#C9D8E8;text-transform:uppercase;'
                         f'letter-spacing:.05em">Harga tanah / m²</div>'
                         f'<div style="font-size:13.5px;font-weight:700;color:#fff">{harga_d}</div></div></div>')

            # Info sections
            if is_s:
                P.append(sh("Identitas"))
                P.append(r2("Pemilik", safe_get(row,"Pemilik"),
                            "Jenis",   safe_get(row,"Jenis_Properti")))
                P.append(r2("Kode inspeksi", safe_get(row,"Kode_Inspeksi"),
                            "Tgl. inspeksi", safe_get(row,"Tanggal_Inspeksi")))
                P.append(r2("Reviewer",      safe_get(row,"Reviewer"),
                            "Pemberi tugas", safe_get(row,"Pemberi_Tugas")))
            else:
                P.append(sh("Properti"))
                thn_v = int(tahun_d) if tahun_d and not pd.isna(tahun_d) else "—"
                P.append(r2("Jenis", safe_get(row,"Jenis_Properti"), "Tahun", thn_v))

            P.append(sh("Lokasi"))
            P.append(r1("Alamat",    safe_get(row,"Alamat")))
            P.append(r1("Kompleks",  safe_get(row,"Kompleks")))
            P.append(r2("Kelurahan", safe_get(row,"Kelurahan"),
                         "Kecamatan", safe_get(row,"Kecamatan")))
            P.append(r2("Kota",     safe_get(row,"Kota"),
                         "Propinsi", safe_get(row,"Propinsi")))

            P.append(sh("Fisik"))
            if is_s:
                # Tentukan apakah ada data bangunan dari sheet terpisah
                kode_ins = safe_get(row, "Kode_Inspeksi")
                bgn_rows = (
                    df_bangunan[df_bangunan["Kode_Inspeksi"] == kode_ins]
                    if not df_bangunan.empty and kode_ins and kode_ins != "—"
                       and "Kode_Inspeksi" in df_bangunan.columns
                    else pd.DataFrame()
                )
                if not bgn_rows.empty:
                    # Ada data dari Data Bangunan — tampilkan Luas Tanah saja di baris pertama
                    P.append(r1("Luas tanah", _luas("Luas_Tanah")))
                    for _, br in bgn_rows.iterrows():
                        jenis_bgn = str(br.get("Jenis_Bangunan", "")) if pd.notna(br.get("Jenis_Bangunan")) else "—"
                        luas_bgn  = br.get("Luas_Bangunan")
                        luas_bgn_fmt = f"{luas_bgn:,.0f} m²" if pd.notna(luas_bgn) else "—"
                        P.append(r2("Jenis bangunan", jenis_bgn, "Luas bangunan", luas_bgn_fmt))
                else:
                    P.append(r2("Luas tanah",    _luas("Luas_Tanah"),
                                 "Luas bangunan", _luas("Luas_Bangunan")))
            else:
                P.append(r2("Luas tanah",    _luas("Luas_Tanah"),
                             "Luas bangunan", _luas("Luas_Bangunan")))
                P.append(r2("Kondisi bangunan", safe_get(row,"Kondisi_Bangunan"),
                            "Kelas bangunan",   safe_get(row,"Kelas_Bangunan")))
            P.append(r2("Peruntukan",  safe_get(row,"Peruntukan"),
                         "Kepemilikan", safe_get(row,"Kepemilikan")))
            if not is_s:
                P.append(sh("Kontak"))
                P.append(r2("Nama", safe_get(row,"Kontak"),
                            "Telepon", safe_get(row,"Telp")))

            P.append('</div>')  # tutup scrollable container

            # Render via components.html agar JavaScript carousel berjalan
            components.html("".join(P), height=640, scrolling=False)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — TABEL DATA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_tabel:
    all_cols = [c for c in filtered.columns if not c.startswith("_")]
    default_cols = [c for c in all_cols if c not in ("Latitude", "Longitude", "Tahun_Bersih")]

    t1, t2, t3 = st.columns([4, 1.3, 1.3], vertical_alignment="bottom")
    cari = t1.text_input("Cari", placeholder="Cari alamat, kelurahan, kecamatan, nomor…",
                         label_visibility="collapsed")
    with t2.popover("Atur kolom", icon=":material/view_column:", use_container_width=True):
        disp_cols = st.multiselect("Kolom yang ditampilkan", all_cols, default=default_cols,
                                   format_func=lambda c: COL_LABELS.get(c, c.replace("_", " ")))

    disp_df = filtered[disp_cols].copy() if disp_cols else pd.DataFrame()
    if cari.strip() and not disp_df.empty:
        _txt = filtered[all_cols].astype(str).agg(" ".join, axis=1).str.lower()
        disp_df = disp_df[_txt.str.contains(re.escape(cari.strip().lower()), na=False)]

    t3.download_button(
        label="Unduh Excel",
        icon=":material/download:",
        data=to_excel_bytes(disp_df),
        file_name=f"data_tanah_{city_label}_{selected_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=disp_df.empty,
    )

    if disp_cols:
        col_cfg = {c: COL_LABELS.get(c, c.replace("_", " ")) for c in disp_cols}
        for c in ("Harga_Tanah", "Harga_Total"):
            if c in disp_cols:
                col_cfg[c] = st.column_config.NumberColumn(COL_LABELS[c], format="localized")
        for c in ("Luas_Tanah", "Luas_Bangunan"):
            if c in disp_cols:
                col_cfg[c] = st.column_config.NumberColumn(COL_LABELS[c], format="localized")
        for c in ("Tahun", "Tahun_Bersih"):
            if c in disp_cols and pd.api.types.is_numeric_dtype(disp_df[c]):
                col_cfg[c] = st.column_config.NumberColumn(COL_LABELS.get(c, c), format="%d")
        view_df = disp_df.copy()
        for c in ("Foto", "Foto_Jalan", "Foto_Dalam", "Foto_Samping_Kanan",
                  "Foto_Samping_Kiri", "Gambar_Situasi"):
            if c in disp_cols:
                view_df[c] = view_df[c].where(view_df[c].astype(str).str.startswith("http"))
                col_cfg[c] = st.column_config.LinkColumn(COL_LABELS.get(c, c), display_text="Buka ↗")

        st.caption(f"Menampilkan {len(disp_df)} dari {len(filtered)} baris.")
        st.dataframe(view_df, use_container_width=True, height=520, hide_index=True,
                     column_config=col_cfg)
    else:
        st.info("Pilih minimal satu kolom melalui tombol **Atur kolom**.")

    outliers_df = filtered[filtered["_outlier"]]
    if not outliers_df.empty:
        with st.expander(f"{len(outliers_df)} data outlier terdeteksi (metode IQR 1,5×)",
                         icon=":material/warning:"):
            show_cols_out = [c for c in ["Nomor", "Alamat", "Kecamatan", "Kota",
                                          "Tahun_Bersih", "Luas_Tanah", "Harga_Tanah"]
                             if c in outliers_df.columns]
            st.dataframe(
                outliers_df[show_cols_out], use_container_width=True, hide_index=True,
                column_config={
                    **{c: COL_LABELS.get(c, c) for c in show_cols_out},
                    "Tahun_Bersih": st.column_config.NumberColumn("Tahun", format="%d"),
                    "Luas_Tanah":   st.column_config.NumberColumn(COL_LABELS["Luas_Tanah"], format="localized"),
                    "Harga_Tanah":  st.column_config.NumberColumn(COL_LABELS["Harga_Tanah"], format="localized"),
                },
            )

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ANALISA PERBANDINGAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_analisa:
    st.caption("Lengkapi data obyek, atur parameter koreksi, lalu pilih data pembanding. "
               "Harga indikasi dan koefisien variasi (CV) dihitung otomatis.")

    # ── Lookup tables ─────────────────────────────────────────────────────────
    _ROAD_SCORE = {"Arteri": 4, "Kolektor": 3, "Lokal": 2, "Lingkungan": 1}
    _LANE_SCORE = {"Gang": 0, "1 lajur": 1, "2 lajur": 2, "4 lajur": 3, "6 lajur": 4, "8 lajur": 5}
    _OWN_OPTS   = ["SHM", "SHSRS", "SHGB", "HGB", "SHP", "HP", "HGU", "Girik/AJB", "Lainnya"]
    _ROAD_OPTS  = list(_ROAD_SCORE.keys())
    _LANE_OPTS  = list(_LANE_SCORE.keys())

    def _road_total(kelas, lajur):
        return _ROAD_SCORE.get(str(kelas), 2) + _LANE_SCORE.get(str(lajur), 1)

    def _peruntukan_score(val):
        """Skor peruntukan: Komersial(5) > Perumahan(4) > Industri(3) > Pertanian(2) > Fasilitias Umum(1)."""
        v = str(val).lower()
        if any(k in v for k in ["komersial", "perdagangan", "jasa", "niaga", "bisnis"]):
            return 5
        if any(k in v for k in ["permukiman", "perumahan", "hunian", "residensial"]):
            return 4
        if "industri" in v:
            return 3
        if any(k in v for k in ["pertanian", "sawah", "kebun", "perkebunan", "ladang"]):
            return 2
        if any(k in v for k in ["fasilitas", "fasum", "sosial", "pendidikan", "kesehatan"]):
            return 1
        return None

    def _detect_kep(val):
        v = str(val).strip().upper().replace(" ", "")
        for k in _OWN_OPTS:
            if k.upper().replace(" ", "") in v:
                return k
        return "Lainnya"

    is_subject_mask = filtered["Nomor"].astype(str).str.lower().str.contains("obyek", na=False)
    subject_rows    = filtered[is_subject_mask]
    comparable_rows = filtered[~is_subject_mask]

    col_subj, col_param = st.columns(2, gap="medium")

    with col_subj:
        with st.container(border=True):
            section_title("Obyek penilaian", 1)
            if not subject_rows.empty:
                s = subject_rows.iloc[0]
                subj_luas  = float(s.get("Luas_Tanah") or 0)
                _h = s.get("Harga_Tanah")
                subj_harga = float(_h) if (_h is not None and not pd.isna(_h)) else 0.0
                harga_disp = format_currency(subj_harga) + "/m²" if subj_harga > 0 else "Belum dinilai"
                _kv = [("Alamat", s.get("Alamat", "-")), ("Kecamatan", s.get("Kecamatan", "-")),
                       ("Luas tanah", f"{subj_luas:,.0f} m²".replace(",", ".")),
                       ("Kepemilikan", s.get("Kepemilikan", "-")), ("Peruntukan", s.get("Peruntukan", "-")),
                       ("Harga awal", harga_disp)]
                st.markdown(
                    '<dl class="kv">' + "".join(
                        f"<dt>{k}</dt><dd>{'—' if pd.isna(v) else v}</dd>" for k, v in _kv
                    ) + "</dl>", unsafe_allow_html=True)
                _kep_def = _detect_kep(s.get("Kepemilikan", "SHM"))
                subj_kep = st.selectbox("Bukti kepemilikan", _OWN_OPTS,
                                        index=_OWN_OPTS.index(_kep_def) if _kep_def in _OWN_OPTS else 0,
                                        key="an_subj_kep")
                # Default tahun dari Tanggal Inspeksi
                _tgl_ins = s.get("Tanggal_Inspeksi")
                try:
                    _default_ref_year = int(pd.to_datetime(str(_tgl_ins), dayfirst=True, errors="coerce").year)
                    if pd.isna(_default_ref_year): _default_ref_year = 2025
                except Exception:
                    _default_ref_year = 2025
                # Skor peruntukan obyek untuk auto-koreksi
                _subj_perun_score = _peruntukan_score(s.get("Peruntukan", ""))
            else:
                st.warning("Baris 'Obyek Penilaian' tidak ditemukan. Isi data obyek secara manual.",
                           icon=":material/edit_note:")
                subj_luas  = st.number_input("Luas tanah obyek (m²)", value=0.0, min_value=0.0, step=10.0)
                subj_harga = 0.0
                subj_kep   = st.selectbox("Bukti kepemilikan", _OWN_OPTS, key="an_subj_kep")
                _default_ref_year = 2025
                _subj_perun_score = None

            jc1, jc2 = st.columns(2)
            subj_road_cls = jc1.selectbox("Kelas jalan", _ROAD_OPTS, index=1, key="an_subj_road_cls")
            subj_lajur    = jc2.selectbox("Jumlah lajur", _LANE_OPTS, index=2, key="an_subj_lajur")

    with col_param:
        with st.container(border=True):
            section_title("Parameter koreksi", 2)
            ref_year = st.number_input("Tahun referensi penilaian", value=_default_ref_year,
                                       min_value=2000, max_value=2100, step=1)
            pc1, pc2 = st.columns(2)
            diskon_pct   = pc1.number_input("Diskon penawaran (%)", value=10.0,
                                            min_value=0.0, max_value=50.0, step=0.5,
                                            help="Diskon dari harga penawaran ke harga transaksi (berlaku untuk semua data pembanding)")
            time_adj_pct = pc2.number_input("Waktu (%/tahun)", value=5.0,
                                            min_value=0.0, max_value=50.0, step=0.5,
                                            help="Kenaikan harga pasar per tahun (positif = pasar naik)")
            size_adj_pct = pc1.number_input("Luas (%/100 m²)", value=0.5,
                                            min_value=0.0, max_value=20.0, step=0.5,
                                            help="Penyesuaian harga akibat perbedaan luas per 100m²")
            lokasi_ppt   = pc2.number_input("Lokasi (%/poin)", value=5.0,
                                            min_value=0.0, max_value=20.0, step=0.5,
                                            help="Penyesuaian per poin perbedaan skor lokasi (kelas jalan + jumlah lajur)")

    comp = None
    with st.container(border=True):
        section_title("Data pembanding & penyesuaian", 3)
        if comparable_rows.empty:
            st.warning("Tidak ada data pembanding tersedia.")
        else:
            nomor_opts = comparable_rows["Nomor"].astype(str).tolist()
            _lbl = dict(zip(
                nomor_opts,
                (comparable_rows["Nomor"].astype(str) + " · " +
                 comparable_rows.get("Alamat", pd.Series("", index=comparable_rows.index))
                 .fillna("").astype(str).str.slice(0, 40)).tolist(),
            ))
            selected   = st.multiselect(
                "Data pembanding (disarankan 3–5 data)",
                options=nomor_opts,
                default=nomor_opts[:min(3, len(nomor_opts))],
                format_func=lambda n: _lbl.get(n, n),
            )

            if not selected:
                st.info("Pilih minimal satu data pembanding untuk memulai perhitungan.")
            else:
                comp = (comparable_rows[comparable_rows["Nomor"].astype(str).isin(selected)]
                        .copy().reset_index(drop=True))

                # Auto-hitung Kor. Peruntukan dari kategori penggunaan tanah
                def _perun_adj(comp_perun_val):
                    comp_score = _peruntukan_score(comp_perun_val)
                    if _subj_perun_score is None or comp_score is None:
                        return 0.0
                    return round((_subj_perun_score - comp_score) * 2.5, 1)

                edit_init = pd.DataFrame({
                    "No":                   comp["Nomor"].astype(str).tolist(),
                    "Alamat":               comp["Alamat"].fillna("—").tolist() if "Alamat" in comp.columns else ["—"] * len(comp),
                    "Kepemilikan":          [
                        _detect_kep(comp.at[i, "Kepemilikan"])
                        if "Kepemilikan" in comp.columns else "Lainnya"
                        for i in range(len(comp))
                    ],
                    "Kelas Jalan":          ["Lokal"]   * len(comp),
                    "Jumlah Lajur":         ["2 lajur"] * len(comp),
                    "Kor. Peruntukan (%)":  [
                        _perun_adj(comp.at[i, "Peruntukan"]) if "Peruntukan" in comp.columns else 0.0
                        for i in range(len(comp))
                    ],
                })

                # Simpan ke session_state agar perubahan pengguna tidak hilang saat rerun
                _ss_adj_key = f"adj_df_{'_'.join(sorted(str(x) for x in selected))}"
                if _ss_adj_key not in st.session_state:
                    st.session_state[_ss_adj_key] = edit_init.copy()

                st.caption("Kepemilikan dan peruntukan terbaca otomatis dari data. "
                           "Kolom berikon pensil dapat diubah langsung di tabel.")
                edited = st.data_editor(
                    st.session_state[_ss_adj_key],
                    use_container_width=True,
                    column_config={
                        "No":                  st.column_config.TextColumn("No", disabled=True, width="small"),
                        "Alamat":              st.column_config.TextColumn("Alamat", disabled=True, width="medium"),
                        "Kepemilikan":         st.column_config.SelectboxColumn("Kepemilikan", options=_OWN_OPTS, width="small"),
                        "Kelas Jalan":         st.column_config.SelectboxColumn("Kelas Jalan", options=_ROAD_OPTS, width="small"),
                        "Jumlah Lajur":        st.column_config.SelectboxColumn("Jumlah Lajur", options=_LANE_OPTS, width="small"),
                        "Kor. Peruntukan (%)": st.column_config.NumberColumn("Kor. Peruntukan", format="%.1f%%",
                                                                               min_value=-50.0, max_value=50.0, step=0.5, width="small"),
                    },
                    hide_index=True,
                )
                # Simpan hasil edit kembali ke session_state
                st.session_state[_ss_adj_key] = edited

                # ── Compute all adjustments ──────────────────────────────
                subj_loc_score = _road_total(subj_road_cls, subj_lajur)

                # Kepemilikan: SHM vs non-SHM = ±5% (flat, bukan per-ranking)
                def _kep_adj(comp_kep):
                    if subj_kep == "SHM" and str(comp_kep) != "SHM":
                        return 5.0   # pembanding lebih rendah → sesuaikan naik
                    elif subj_kep != "SHM" and str(comp_kep) == "SHM":
                        return -5.0  # pembanding lebih tinggi → sesuaikan turun
                    return 0.0

                comp["Kor_Kepemilikan_%"] = edited["Kepemilikan"].apply(_kep_adj)
                comp["Kor_Lokasi_%"] = edited.apply(
                    lambda r: (subj_loc_score - _road_total(r["Kelas Jalan"], r["Jumlah Lajur"])) * lokasi_ppt,
                    axis=1,
                )
                comp["Kor_Peruntukan_%"] = pd.to_numeric(
                    edited["Kor. Peruntukan (%)"], errors="coerce"
                ).fillna(0.0).values

                comp["Selisih_Tahun"]  = ref_year - comp["Tahun_Bersih"].fillna(ref_year)
                comp["Kor_Waktu_%"]    = comp["Selisih_Tahun"] * time_adj_pct
                if subj_luas > 0:
                    comp["Selisih_Luas"] = comp["Luas_Tanah"].fillna(subj_luas) - subj_luas
                    comp["Kor_Luas_%"]   = -(comp["Selisih_Luas"] / 100) * size_adj_pct
                else:
                    comp["Kor_Luas_%"]   = 0.0

                # ── Total & bobot penyesuaian ────────────────────────────
                comp["Kor_Diskon_%"] = -diskon_pct
                _kor_cols = ["Kor_Diskon_%", "Kor_Waktu_%", "Kor_Luas_%",
                             "Kor_Kepemilikan_%", "Kor_Lokasi_%", "Kor_Peruntukan_%"]
                comp["Total_Penyesuaian_%"] = comp[_kor_cols].sum(axis=1).round(2)
                comp["Total_Absolut_%"]     = comp[_kor_cols].abs().sum(axis=1).round(2)
                # Bobot: berbanding terbalik dengan total absolut
                _inv = 1.0 / (comp["Total_Absolut_%"] + 0.01)
                comp["Bobot_%"] = (_inv / _inv.sum() * 100).round(1)
                # Harga Final = Harga Awal × (1 + Total Penyesuaian) × Bobot
                comp["Harga_Stl_Koreksi"] = comp["Harga_Tanah"] * (1 + comp["Total_Penyesuaian_%"] / 100)
                comp["Harga_Final"]        = comp["Harga_Stl_Koreksi"] * comp["Bobot_%"] / 100

    # ── Hasil analisa (lebar penuh) ──────────────────────────────────────────
    if comp is not None:
        tbl = comp[[
            "Nomor", "Alamat", "Tahun_Bersih", "Luas_Tanah",
            "Harga_Tanah",
            "Kor_Diskon_%", "Kor_Waktu_%", "Kor_Luas_%",
            "Kor_Kepemilikan_%", "Kor_Lokasi_%", "Kor_Peruntukan_%",
            "Total_Penyesuaian_%", "Total_Absolut_%",
            "Bobot_%",
            "Harga_Stl_Koreksi",
            "Harga_Final",
        ]].copy()
        tbl.columns = [
            "Nomor", "Alamat", "Tahun", "Luas (m²)",
            "Harga Awal",
            "Kor. Diskon (%)", "Kor. Waktu (%)", "Kor. Luas (%)",
            "Kor. Kepemilikan (%)", "Kor. Lokasi (%)", "Kor. Peruntukan (%)",
            "Total Penyesuaian (%)", "Total Absolut (%)",
            "Bobot (%)",
            "Harga Terkoreksi",
            "Harga Final",
        ]

        # Harga Final sudah menyertakan bobot, jadi indikasi = sum(Harga_Final)
        harga_indikasi = comp["Harga_Final"].sum()
        cv = (
            comp["Harga_Stl_Koreksi"].std() / comp["Harga_Stl_Koreksi"].mean() * 100
            if len(comp) > 1 and comp["Harga_Stl_Koreksi"].mean() != 0
            else 0.0
        )
        if cv <= 20:
            cv_badge, cv_note = '<span class="badge ok">Homogen</span>', "Hasil dapat diandalkan."
        elif cv <= 30:
            cv_badge, cv_note = '<span class="badge warn">Cukup beragam</span>', "Pertimbangkan meninjau ulang data pembanding."
        else:
            cv_badge, cv_note = '<span class="badge bad">Sangat beragam</span>', "Ganti atau perbaiki data pembanding."

        with st.container(border=True):
            section_title("Hasil analisa", 4)
            _items = [
                ("Harga indikasi", rp(harga_indikasi), "Rata-rata tertimbang harga terkoreksi", True),
                ("Koefisien variasi", f"{cv:.1f}% {cv_badge}", cv_note, False),
                ("Jumlah pembanding", f"{len(comp)}", "Disarankan 3–5 data", False),
            ]
            if subj_harga > 0:
                _sel = (harga_indikasi - subj_harga) / subj_harga * 100
                _items.append(("Harga awal obyek", rp(subj_harga), f"Selisih indikasi {_sel:+.1f}%", False))
            kpi_cards(_items)

            fig = go.Figure()
            fig.add_trace(go.Bar(x=comp["Nomor"].astype(str), y=comp["Harga_Tanah"],
                                 name="Harga awal", marker_color="#C9D8E8",
                                 hovertemplate="Data %{x}<br>Rp %{y:,.0f}/m²<extra>Harga awal</extra>"))
            fig.add_trace(go.Bar(x=comp["Nomor"].astype(str), y=comp["Harga_Stl_Koreksi"],
                                 name="Harga terkoreksi", marker_color=NAVY,
                                 hovertemplate="Data %{x}<br>Rp %{y:,.0f}/m²<extra>Harga terkoreksi</extra>"))
            if subj_harga > 0:
                fig.add_hline(y=subj_harga, line_dash="dash", line_color=RED, line_width=1.5,
                              annotation_text=f"Harga obyek {format_currency(subj_harga)}",
                              annotation_font_color=RED, annotation_position="top left")
            fig.add_hline(y=harga_indikasi, line_dash="dot", line_color=TEAL, line_width=2,
                          annotation_text=f"Indikasi {format_currency(harga_indikasi)}",
                          annotation_font_color=TEAL, annotation_position="top right")
            fig.update_layout(barmode="group", bargap=0.3, xaxis_title="Nomor data",
                              yaxis_title="Harga (Rp/m²)")
            fig.update_xaxes(type="category")
            st.plotly_chart(style_fig(fig, "Harga awal dan terkoreksi per data pembanding", height=360),
                            use_container_width=True, config=PLOT_CFG)

            _pct = st.column_config.NumberColumn(format="%.1f%%")
            st.dataframe(
                tbl,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Tahun":                   st.column_config.NumberColumn(format="%d"),
                    "Harga Awal":              st.column_config.NumberColumn("Harga Awal (Rp/m²)",       format="localized"),
                    "Harga Terkoreksi":        st.column_config.NumberColumn("Harga Terkoreksi (Rp/m²)", format="localized"),
                    "Harga Final":             st.column_config.NumberColumn("Harga Final (Rp/m²)",      format="localized"),
                    "Kor. Diskon (%)":         _pct,
                    "Kor. Waktu (%)":          _pct,
                    "Kor. Luas (%)":           _pct,
                    "Kor. Kepemilikan (%)":    _pct,
                    "Kor. Lokasi (%)":         _pct,
                    "Kor. Peruntukan (%)":     _pct,
                    "Total Penyesuaian (%)":   st.column_config.NumberColumn(format="%.2f%%"),
                    "Total Absolut (%)":       st.column_config.NumberColumn(format="%.2f%%"),
                    "Bobot (%)":               _pct,
                    "Luas (m²)":               st.column_config.NumberColumn(format="localized"),
                },
            )

            st.download_button(
                label="Unduh hasil analisa (Excel)",
                icon=":material/download:",
                data=to_excel_bytes(tbl),
                file_name=f"analisa_perbandingan_{city_label}_{selected_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
