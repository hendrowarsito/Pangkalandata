import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster, HeatMap
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
    page_icon="🏡"
)

st.markdown("""
<style>
.block-container { padding-top: 1rem; }
.stMetric { background: #f8f9fa; border-radius: 8px; padding: 0.5rem; border-left: 4px solid #667eea; }
.legend-box { position:fixed; bottom:30px; left:30px; z-index:1000; background:white;
              padding:10px 14px; border-radius:8px; border:2px solid #ccc;
              font-size:12px; font-family:sans-serif; box-shadow: 2px 2px 6px rgba(0,0,0,0.2); }
</style>
""", unsafe_allow_html=True)

YEAR_COLORS = {
    "gte_2025": "green",
    "gte_2024": "blue",
    "gte_2023": "orange",
    "lt_2023":  "red",
    "subject":  "purple",
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

def build_foto_html(foto_url):
    """Thumbnail inline + tombol buka tab baru."""
    has_link = bool(foto_url) and str(foto_url).strip() not in ("", "#", "nan", "None", "-")
    if not has_link:
        return ""

    foto_url = str(foto_url).strip()
    lh3, thumb = gdrive_thumbnail(foto_url, width=280)

    if lh3:
        # Coba lh3 dulu; jika gagal, otomatis fallback ke thumbnail API via onerror
        return f"""
        <div style="margin:6px 0;text-align:center">
          <img src="{lh3}"
               style="width:100%;max-width:280px;border-radius:6px;
                      border:1px solid #ddd;display:block;margin:0 auto 4px"
               referrerpolicy="no-referrer"
               onerror="this.src='{thumb}';this.onerror=null;"
          >
          <a href="{foto_url}" target="_blank"
             style="font-size:11px;color:#2980b9;text-decoration:none">
            &#128247; Buka foto selengkapnya &#8599;
          </a>
        </div>"""

    # Bukan URL Google Drive — tampilkan sebagai link biasa
    return (f'<a href="{foto_url}" target="_blank" '
            f'style="font-size:12px">&#128247; Lihat Foto</a><br>')

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
st.sidebar.markdown("## 🏡 Pangkalan Data Tanah\n**KJPP Suwendho Rinaldy dan Rekan**")
st.sidebar.divider()
st.sidebar.header("🔧 Filter Data")
file = st.sidebar.file_uploader("📂 Unggah file Excel data tanah", type=["xlsx"])

if file and "last_file" in st.session_state and file != st.session_state["last_file"]:
    st.session_state["tampilkan"] = False
st.session_state["last_file"] = file

if not file:
    st.markdown("""
    ## 🏡 Pangkalan Data Penilaian Tanah
    ### KJPP Suwendho Rinaldy dan Rekan

    Aplikasi ini mendukung proses **reviu penilaian tanah** dengan fitur:

    | Fitur | Keterangan |
    |---|---|
    | 📊 Dashboard Analitik | Statistik ringkasan, distribusi harga, tren per tahun |
    | 🗺️ Peta Interaktif | Semua data pembanding terpetakan + heatmap harga |
    | 📋 Tabel Data | Filter lanjutan, deteksi outlier, ekspor Excel |
    | 🔄 Analisa Perbandingan | Koreksi waktu & luas, hitung CV, indikasi nilai |

    **Mulai dengan mengunggah file Excel data tanah di sidebar kiri.**
    """)
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
    for col in ["Harga_Total", "Luas_Tanah", "Luas_Bangunan"]:
        if col in df.columns:
            df[col] = df[col].apply(parse_indo_number)

    # Harga per m²
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
df["Tahun_Bersih"] = df["Tahun"].apply(bersihkan_tahun) if "Tahun" in df.columns else pd.Series(dtype=float)

# Tunjukkan format yang terdeteksi
_fmt = df["_format"].iloc[0] if "_fmt" not in st.session_state and not df.empty and "_format" in df.columns else ""
if _fmt == "multi-sheet":
    n_prop = int((df["_sumber"] == "Data Properti").sum())   if "_sumber" in df.columns else 0
    n_pemb = int((df["_sumber"] == "Data Pembanding").sum()) if "_sumber" in df.columns else 0
    st.sidebar.success(f"✅ Format multi-sheet: {n_prop} Obyek Penilaian + {n_pemb} Data Pembanding")

# ─── Filter controls ──────────────────────────────────────────────────────────
city_input = st.sidebar.text_input("🔍 Cari Kota/Kabupaten (sebagian nama OK):")

available_years = sorted([int(y) for y in df["Tahun_Bersih"].dropna().unique()], reverse=True)
selected_year = st.sidebar.selectbox("📅 Pilih Tahun Data:", ["Semua Tahun"] + [str(y) for y in available_years])

kec_options = ["Semua Kecamatan"] + sorted(df["Kecamatan"].dropna().astype(str).unique().tolist())
selected_kecamatan = st.sidebar.selectbox("🏘️ Pilih Kecamatan:", kec_options)

harga_col = df["Harga_Tanah"].dropna()
if not harga_col.empty:
    h_min, h_max = float(harga_col.min()), float(harga_col.max())
    price_range = st.sidebar.slider("💰 Rentang Harga (Rp/m²):", h_min, h_max, (h_min, h_max), format="%.0f")
else:
    price_range = None

luas_col = df["Luas_Tanah"].dropna()
if not luas_col.empty:
    l_min, l_max = float(luas_col.min()), float(luas_col.max())
    luas_range = st.sidebar.slider("📐 Rentang Luas Tanah (m²):", l_min, l_max, (l_min, l_max), format="%.0f")
else:
    luas_range = None

st.sidebar.divider()
st.sidebar.markdown("**Opsi Peta:**")
show_heatmap  = st.sidebar.checkbox("🌡️ Heatmap Harga", value=False)
use_clustering = st.sidebar.checkbox("🔵 Marker Clustering", value=True)

if "tampilkan" not in st.session_state:
    st.session_state["tampilkan"] = False
if st.sidebar.button("🔍 Tampilkan Data", use_container_width=True, type="primary"):
    st.session_state["tampilkan"] = True

if not st.session_state["tampilkan"]:
    st.sidebar.info("Tekan tombol di atas untuk melihat hasil.")
    st.stop()

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
st.markdown(f"### Hasil Filter: **{len(filtered)} data** — Kota: *{city_label}* | Tahun: *{selected_year}*")

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_dashboard, tab_peta, tab_tabel, tab_analisa = st.tabs([
    "📊 Dashboard Analitik",
    "🗺️ Peta Lokasi",
    "📋 Tabel Data",
    "🔄 Analisa Perbandingan",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
with tab_dashboard:
    if filtered.empty:
        st.warning("Tidak ada data yang sesuai dengan filter yang dipilih.")
        st.stop()

    prices = filtered["Harga_Tanah"].dropna()
    outlier_count = int(filtered["_outlier"].sum())

    # KPI row
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 Total Data",       f"{len(filtered):,}")
    k2.metric("💰 Rata-rata",         format_currency(prices.mean()) if not prices.empty else "N/A")
    k3.metric("📉 Minimum",           format_currency(prices.min())  if not prices.empty else "N/A")
    k4.metric("📈 Maksimum",          format_currency(prices.max())  if not prices.empty else "N/A")
    k5.metric("📊 Median",            format_currency(prices.median()) if not prices.empty else "N/A")

    if outlier_count:
        st.warning(
            f"⚠️ **{outlier_count} data outlier** terdeteksi (metode IQR 1.5×). "
            "Detail tersedia di tab **Tabel Data**."
        )

    st.divider()
    col_l, col_r = st.columns(2)

    with col_l:
        # Distribusi harga
        if not prices.empty:
            fig_hist = px.histogram(
                filtered[filtered["Harga_Tanah"].notna()],
                x="Harga_Tanah", nbins=25,
                title="Distribusi Harga Tanah (Rp/m²)",
                labels={"Harga_Tanah": "Harga (Rp/m²)", "count": "Jumlah"},
                color_discrete_sequence=["#667eea"],
            )
            fig_hist.update_layout(showlegend=False, margin=dict(t=40, b=20))
            st.plotly_chart(fig_hist, use_container_width=True)

        # Scatter harga vs luas
        scatter_df = filtered[filtered["Harga_Tanah"].notna() & filtered["Luas_Tanah"].notna()].copy()
        scatter_df["Tahun_Label"] = scatter_df["Tahun_Bersih"].astype(str)
        if not scatter_df.empty:
            fig_scatter = px.scatter(
                scatter_df, x="Luas_Tanah", y="Harga_Tanah",
                color="Tahun_Label",
                hover_data=["Nomor", "Alamat", "Kecamatan"],
                title="Harga vs Luas Tanah",
                labels={"Luas_Tanah": "Luas Tanah (m²)", "Harga_Tanah": "Harga (Rp/m²)", "Tahun_Label": "Tahun"},
                trendline="ols",
            )
            fig_scatter.update_layout(margin=dict(t=40, b=20))
            st.plotly_chart(fig_scatter, use_container_width=True)

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
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Scatter(
                x=by_year["Tahun"], y=by_year["rata_rata"],
                mode="lines+markers", name="Rata-rata",
                line=dict(color="#667eea", width=3), marker=dict(size=8),
            ))
            fig_trend.add_trace(go.Scatter(
                x=by_year["Tahun"], y=by_year["maksimum"],
                mode="lines", name="Maksimum", line=dict(color="#2ecc71", dash="dash"),
            ))
            fig_trend.add_trace(go.Scatter(
                x=by_year["Tahun"], y=by_year["minimum"],
                mode="lines", name="Minimum", line=dict(color="#e74c3c", dash="dash"),
            ))
            fig_trend.update_layout(
                title="Tren Harga per Tahun",
                xaxis_title="Tahun", yaxis_title="Harga (Rp/m²)",
                hovermode="x unified", margin=dict(t=40, b=20),
            )
            st.plotly_chart(fig_trend, use_container_width=True)

        # Rata-rata harga per kecamatan
        by_kec = (
            filtered.groupby("Kecamatan")["Harga_Tanah"]
            .mean()
            .dropna()
            .sort_values(ascending=True)
            .reset_index()
        )
        if not by_kec.empty:
            fig_kec = px.bar(
                by_kec, x="Harga_Tanah", y="Kecamatan", orientation="h",
                title="Rata-rata Harga per Kecamatan",
                labels={"Harga_Tanah": "Rata-rata Harga (Rp/m²)", "Kecamatan": ""},
                color="Harga_Tanah", color_continuous_scale="Viridis",
            )
            fig_kec.update_layout(
                showlegend=False, coloraxis_showscale=False, margin=dict(t=40, b=20)
            )
            st.plotly_chart(fig_kec, use_container_width=True)

    # Ringkasan statistik
    st.subheader("📊 Ringkasan Statistik")
    num_cols = [c for c in ["Harga_Tanah", "Luas_Tanah", "Luas_Bangunan"] if c in filtered.columns]
    if num_cols:
        stats = filtered[num_cols].describe()
        stats.index = ["Jumlah", "Rata-rata", "Std Deviasi", "Minimum",
                       "Kuartil-1", "Median", "Kuartil-3", "Maksimum"]
        st.dataframe(stats.style.format("{:,.2f}"), use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PETA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_peta:
    if filtered.empty:
        st.warning("Tidak ada data untuk ditampilkan di peta.")
    else:
        map_df = filtered[filtered["Latitude"].notna() & filtered["Longitude"].notna()].copy()
        map_df = map_df[
            map_df["Latitude"].between(-90, 90) &
            map_df["Longitude"].between(-180, 180)
        ]

        subj_mask   = map_df["Nomor"].astype(str).str.lower().str.contains("obyek", na=False)
        subj_df     = map_df[subj_mask]
        comp_map_df = map_df[~subj_mask]

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
        )

        folium.TileLayer("OpenStreetMap", name="OpenStreetMap").add_to(m)
        folium.TileLayer(
            tiles="https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png",
            name="Voyager", attr="©OpenStreetMap ©CartoDB",
        ).add_to(m)
        folium.TileLayer(
            tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
            name="Satellite", attr="Tiles © Esri",
        ).add_to(m)

        if show_heatmap and not comp_map_df.empty:
            heat = [
                [r.Latitude, r.Longitude, float(r.Harga_Tanah)]
                for r in comp_map_df.itertuples()
                if pd.notna(getattr(r, "Harga_Tanah", None))
            ]
            if heat:
                HeatMap(heat, name="Heatmap Harga", radius=30, blur=20, min_opacity=0.4).add_to(m)

        # ── Garis jarak dari Obyek Penilaian ke setiap Data Pembanding ──────
        if not subj_df.empty and not comp_map_df.empty:
            s = subj_df.iloc[0]
            lines_fg = folium.FeatureGroup(name="📏 Garis Jarak", show=True).add_to(m)

            for r in comp_map_df.itertuples():
                nomor_r = str(safe_get(r, "Nomor")).strip()
                dist_km = haversine_km(s.Latitude, s.Longitude, r.Latitude, r.Longitude)
                dist_label = f"{dist_km:.2f} km" if dist_km >= 1 else f"{dist_km*1000:.0f} m"

                folium.PolyLine(
                    locations=[[s.Latitude, s.Longitude], [r.Latitude, r.Longitude]],
                    color="#e74c3c",
                    weight=2,
                    dash_array="6 4",
                    opacity=0.8,
                    tooltip=f"Obyek → Data {nomor_r}: {dist_label}",
                ).add_to(lines_fg)

                # Label jarak di tengah garis
                mid_lat = (s.Latitude + r.Latitude) / 2
                mid_lon = (s.Longitude + r.Longitude) / 2
                folium.Marker(
                    [mid_lat, mid_lon],
                    icon=folium.DivIcon(
                        html=f"""
                        <div style="font-size:10px;font-weight:bold;color:#c0392b;
                                    background:rgba(255,255,255,0.9);padding:2px 6px;
                                    border-radius:10px;border:1px solid #e74c3c;
                                    white-space:nowrap;pointer-events:none;
                                    box-shadow:1px 1px 3px rgba(0,0,0,0.2)">
                            📏 {dist_label}
                        </div>""",
                        icon_size=(90, 22),
                        icon_anchor=(45, 11),
                    ),
                ).add_to(lines_fg)

        # ── Layer marker data pembanding (bisa di-cluster) ──────────────────
        if use_clustering:
            marker_layer = MarkerCluster(name="Data Pembanding").add_to(m)
        else:
            marker_layer = folium.FeatureGroup(name="Data Pembanding").add_to(m)

        # ── Layer Obyek Penilaian — selalu terpisah, tidak masuk cluster ────
        subj_layer = folium.FeatureGroup(name="🏠 Obyek Penilaian", show=True).add_to(m)

        def _foto_mini(url, label_txt):
            """Thumbnail kecil dengan label dan link buka tab baru."""
            lh3, thumb = gdrive_thumbnail(url, width=130)
            if not lh3:
                return ""
            return f"""
            <div style="display:inline-block;text-align:center;margin:3px 3px 0 0;vertical-align:top">
              <img src="{lh3}"
                   style="width:130px;height:90px;object-fit:cover;border-radius:4px;
                          border:1px solid #ddd"
                   referrerpolicy="no-referrer"
                   onerror="this.src='{thumb}';this.onerror=null;">
              <br>
              <a href="{url}" target="_blank"
                 style="font-size:10px;color:#2980b9">&#8599; {label_txt}</a>
            </div>"""

        # ── CSS constants untuk popup ─────────────────────────────────────
        SEC  = ("background:#f0faf5;color:#1a7a4a;font-size:10px;font-weight:bold;"
                "text-transform:uppercase;letter-spacing:0.6px;padding:3px 7px;"
                "border-radius:3px;margin:8px 0 3px 0;display:block")
        TD_L = ("color:#777;white-space:nowrap;padding:3px 6px 3px 0;"
                "vertical-align:top;font-size:11.5px")
        TD_V = "padding:3px 10px 3px 0;vertical-align:top;font-size:12px"
        TR_A = "background:#f8fafb"   # alternating row background

        def build_popup(r, is_subj, tahun, harga_fmt, luas_t, luas_b, foto):
            nomor_str = str(safe_get(r, "Nomor")).strip()

            # ── row helpers ───────────────────────────────────────────────
            def r2(l1, v1, l2, v2, stripe=False):
                bg = f' style="background:#f8fafb"' if stripe else ""
                return (
                    f'<tr{bg}>'
                    f'<td style="{TD_L}">{l1}</td>'
                    f'<td style="{TD_V}">{v1}</td>'
                    f'<td style="{TD_L}">{l2}</td>'
                    f'<td style="{TD_V}">{v2}</td>'
                    f'</tr>'
                )
            def r1(lbl, val, stripe=False):
                bg = f' style="background:#f8fafb"' if stripe else ""
                return (
                    f'<tr{bg}>'
                    f'<td style="{TD_L}">{lbl}</td>'
                    f'<td colspan="3" style="{TD_V}">{val}</td>'
                    f'</tr>'
                )
            def sec(icon, title):
                return f'<tr><td colspan="4"><span style="{SEC}">{icon} {title}</span></td></tr>'

            tbl_open  = '<table style="width:100%;border-collapse:collapse">'
            tbl_close = '</table>'

            if is_subj:
                # ── Galeri foto ───────────────────────────────────────────
                foto_pairs = [
                    (safe_get(r, "Foto",              "#"), "Depan"),
                    (safe_get(r, "Foto_Dalam",         "#"), "Dalam"),
                    (safe_get(r, "Foto_Samping_Kanan", "#"), "Kanan"),
                    (safe_get(r, "Foto_Samping_Kiri",  "#"), "Kiri"),
                    (safe_get(r, "Gambar_Situasi",     "#"), "Situasi"),
                ]
                thumbs = "".join(_foto_mini(u, lb) for u, lb in foto_pairs
                                 if u not in ("#", "-", "nan", "None", ""))
                galeri = (
                    f'<div style="margin:6px 0 10px;overflow-x:auto;white-space:nowrap">{thumbs}</div>'
                    if thumbs else ""
                )
                sv_url = generate_streetview_url(r.Latitude, r.Longitude)
                body = f"""
                {galeri}
                {tbl_open}
                  {sec('📋','Identitas')}
                  {r2('Pemilik', safe_get(r,'Pemilik'), 'Jenis', safe_get(r,'Jenis_Properti'))}
                  {r2('Kode Inspeksi', safe_get(r,'Kode_Inspeksi'), 'Reviewer', safe_get(r,'Reviewer'), True)}
                  {r1('Pemberi Tugas', safe_get(r,'Pemberi_Tugas'))}
                  {sec('📍','Lokasi')}
                  {r1('Alamat', safe_get(r,'Alamat'))}
                  {r2('Kelurahan', safe_get(r,'Kelurahan'), 'Kecamatan', safe_get(r,'Kecamatan'), True)}
                  {r2('Kota', safe_get(r,'Kota'), 'Propinsi', safe_get(r,'Propinsi'))}
                  {sec('📐','Fisik')}
                  {r2('Luas Tanah', f'{luas_t} m²', 'Luas Bangunan', f'{luas_b} m²')}
                  {r2('Peruntukan', safe_get(r,'Peruntukan'), 'Kepemilikan', safe_get(r,'Kepemilikan'), True)}
                  {r1('Penggunaan', safe_get(r,'Penggunaan'))}
                {tbl_close}
                <div style="margin-top:8px;font-size:11px;display:flex;gap:14px">
                  <a href="{sv_url}" target="_blank"
                     style="color:#2980b9;text-decoration:none">
                    &#128269; Street View &#8599;
                  </a>
                </div>"""

                return f"""
                <div style="font-family:'Segoe UI',Arial,sans-serif;
                            min-width:320px;max-width:380px">
                  <div style="background:#fdecea;border-left:4px solid #c0392b;
                              padding:7px 10px;margin-bottom:4px;
                              border-radius:0 5px 5px 0">
                    <b style="font-size:14px;color:#c0392b">🏠 Obyek Penilaian</b>
                  </div>
                  {body}
                </div>"""

            else:
                # ── Data Pembanding ───────────────────────────────────────
                foto_jalan_url = str(safe_get(r, "Foto_Jalan", "#"))
                foto_jalan_ok  = foto_jalan_url not in ("#", "-", "nan", "None", "")
                sv_url         = generate_streetview_url(r.Latitude, r.Longitude)
                lh3, thumb_url = gdrive_thumbnail(foto, width=440)

                # ── foto + semua link dalam SATU baris ────────────────────
                links = []
                if lh3 and foto not in ("#", "-", "nan", "None", ""):
                    links.append(
                        f'<a href="{foto}" target="_blank"'
                        f' style="color:#2980b9;text-decoration:none">&#128247; Foto Depan &#8599;</a>'
                    )
                if foto_jalan_ok:
                    links.append(
                        f'<a href="{foto_jalan_url}" target="_blank"'
                        f' style="color:#2980b9;text-decoration:none">&#128247; Foto Jalan &#8599;</a>'
                    )
                links.append(
                    f'<a href="{sv_url}" target="_blank"'
                    f' style="color:#2980b9;text-decoration:none">&#128269; Street View &#8599;</a>'
                )
                link_bar = (
                    f'<div style="display:flex;gap:12px;margin:5px 0 10px;'
                    f'font-size:11px;flex-wrap:wrap">'
                    + " ".join(links) +
                    f'</div>'
                )

                foto_img = ""
                if lh3:
                    foto_img = f"""
                    <img src="{lh3}"
                         style="width:100%;border-radius:7px;display:block;
                                border:1px solid #e0e0e0;
                                box-shadow:0 2px 8px rgba(0,0,0,0.12)"
                         referrerpolicy="no-referrer"
                         onerror="this.src='{thumb_url}';this.onerror=null;">"""

                # Harga
                ht  = getattr(r, "Harga_Total", None)
                ht_str = format_currency(ht) if (ht and not pd.isna(ht)) else None

                harga_card = f"""
                <div style="background:linear-gradient(135deg,#eafaf1,#d5f5e3);
                            border-radius:7px;padding:8px 12px;margin:6px 0;
                            border:1px solid #a9dfbf">
                  {"<div style='font-size:11px;color:#555;margin-bottom:2px'>Harga Total: " + ht_str + "</div>" if ht_str else ""}
                  <div style="font-size:22px;font-weight:bold;color:#1a7a4a;line-height:1.1">
                    {harga_fmt}
                    <span style="font-size:12px;color:#555;font-weight:normal">/m²</span>
                  </div>
                </div>"""

                # Badge header
                jenis_data = safe_get(r, "Jenis_Data")
                badge_col  = "#e67e22" if "penawaran" in str(jenis_data).lower() else "#2980b9"
                header = f"""
                <div style="background:#eafaf1;border-left:4px solid #27ae60;
                            padding:7px 10px;margin-bottom:6px;border-radius:0 5px 5px 0;
                            display:flex;align-items:center;gap:8px">
                  <b style="font-size:15px;color:#1a7a4a">Data {nomor_str}</b>
                  <span style="background:{badge_col};color:white;border-radius:10px;
                               padding:1px 8px;font-size:10.5px;font-weight:600">
                    {jenis_data}
                  </span>
                  <span style="color:#888;font-size:11px;margin-left:auto">{int(tahun) if tahun else ''}</span>
                </div>"""

                body = f"""
                {foto_img}
                {link_bar}
                {harga_card}
                {tbl_open}
                  {sec('📍','Lokasi & Properti')}
                  {r2('Jenis', safe_get(r,'Jenis_Properti'), 'Tahun', int(tahun) if tahun else '-')}
                  {r1('Alamat', safe_get(r,'Alamat'), True)}
                  {r1('Kompleks', safe_get(r,'Kompleks'))}
                  {r2('Kelurahan', safe_get(r,'Kelurahan'), 'Kecamatan', safe_get(r,'Kecamatan'), True)}
                  {r2('Kota', safe_get(r,'Kota'), 'Propinsi', safe_get(r,'Propinsi'))}
                  {sec('🏗️','Fisik Bangunan')}
                  {r2('Luas Tanah', f'{luas_t} m²', 'Luas Bangunan', f'{luas_b} m²')}
                  {r2('Kondisi Bgn', safe_get(r,'Kondisi_Bangunan'), 'Kelas Bgn', safe_get(r,'Kelas_Bangunan'), True)}
                  {sec('📞','Kontak')}
                  {r2('Nama', safe_get(r,'Kontak'), 'Telp', safe_get(r,'Telp'))}
                {tbl_close}"""

                return f"""
                <div style="font-family:'Segoe UI',Arial,sans-serif;
                            min-width:440px;max-width:500px">
                  {header}
                  {body}
                </div>"""

        for r in map_df.itertuples():
            nomor     = str(safe_get(r, "Nomor")).strip()
            tahun     = getattr(r, "Tahun_Bersih", 0) or 0
            is_subj   = "obyek" in nomor.lower()
            foto      = str(safe_get(r, "Foto", "#"))
            harga_fmt = format_currency(getattr(r, "Harga_Tanah", 0))
            def _fmt_luas(col):
                v = getattr(r, col, None)
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return "-"
                try:
                    return f"{float(v):,.0f}".replace(",", ".")
                except Exception:
                    return str(v)
            luas_t = _fmt_luas("Luas_Tanah")
            luas_b = _fmt_luas("Luas_Bangunan")

            popup_html = build_popup(r, is_subj, tahun, harga_fmt, luas_t, luas_b, foto)
            target = subj_layer if is_subj else marker_layer

            if is_subj:
                # Pin merah besar berbentuk drop-pin dengan ikon rumah — mudah dibedakan
                folium.Marker(
                    location=[r.Latitude, r.Longitude],
                    popup=folium.Popup(popup_html, max_width=380),
                    tooltip="🏠 Obyek Penilaian — klik untuk detail",
                    icon=folium.DivIcon(
                        html="""
                        <div style="position:relative;width:38px;height:50px">
                          <div style="width:38px;height:38px;background:#c0392b;
                                      border-radius:50% 50% 50% 0;
                                      transform:rotate(-45deg);
                                      border:3px solid white;
                                      box-shadow:0 3px 8px rgba(0,0,0,0.5)">
                          </div>
                          <span style="position:absolute;top:4px;left:6px;
                                       font-size:18px;line-height:1">🏠</span>
                        </div>""",
                        icon_size=(38, 50),
                        icon_anchor=(19, 50),
                    ),
                ).add_to(target)
                # Label Obyek Penilaian — lebih besar & mencolok
                folium.Marker(
                    [r.Latitude, r.Longitude],
                    icon=folium.DivIcon(
                        html="""
                        <div style="font-size:12px;font-weight:bold;color:#c0392b;
                                    background:rgba(255,255,255,0.95);padding:3px 8px;
                                    border-radius:5px;white-space:nowrap;
                                    border:2px solid #c0392b;pointer-events:none;
                                    box-shadow:1px 2px 5px rgba(0,0,0,0.3);margin-top:2px">
                            🏠 OBYEK PENILAIAN
                        </div>""",
                        icon_size=(175, 28),
                        icon_anchor=(87, -4),
                    ),
                ).add_to(m)
            else:
                warna = get_color_by_year(tahun)
                folium.Marker(
                    location=[r.Latitude, r.Longitude],
                    popup=folium.Popup(popup_html, max_width=500),
                    tooltip=f"Data {nomor} | {harga_fmt}/m²",
                    icon=folium.Icon(color=warna, icon="info-sign", prefix="glyphicon"),
                ).add_to(target)

                label_color = "#922b21" if warna == "red" else "#154360"
                folium.Marker(
                    [r.Latitude, r.Longitude],
                    icon=folium.DivIcon(
                        html=f"""
                        <div style="font-size:11px;color:{label_color};font-weight:bold;
                                    background:rgba(255,255,255,0.85);padding:2px 5px;
                                    border-radius:4px;white-space:nowrap;
                                    border:1px solid {label_color};pointer-events:none">
                            {harga_fmt}/m²<br>
                            <span style="font-size:10px">Data {nomor}</span>
                        </div>""",
                        icon_size=(140, 36),
                        icon_anchor=(0, 0),
                    ),
                ).add_to(m)

        folium.LayerControl(collapsed=False).add_to(m)

        legend = """
        <div class="legend-box">
          <b>Legenda Tahun:</b><br>
          <span style="color:green">&#9679;</span> &ge; 2025<br>
          <span style="color:blue">&#9679;</span> 2024<br>
          <span style="color:orange">&#9679;</span> 2023<br>
          <span style="color:red">&#9679;</span> &lt; 2023<br>
          <span style="color:#c0392b;font-size:14px">🏠</span> Obyek Penilaian<br>
          <span style="color:#e74c3c">&#9135;&#9135;</span> Garis Jarak
        </div>
        """
        m.get_root().html.add_child(folium.Element(legend))

        n_comp = len(comp_map_df)
        n_subj = len(subj_df)
        st.caption(
            f"Menampilkan **{n_subj} Obyek Penilaian** + **{n_comp} Data Pembanding** "
            f"({'dengan' if not subj_df.empty else 'tanpa'} garis jarak)"
        )
        st_folium(m, width="100%", height=1000, returned_objects=[])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — TABEL DATA
# ═══════════════════════════════════════════════════════════════════════════════
with tab_tabel:
    if filtered.empty:
        st.warning("Tidak ada data yang sesuai dengan filter.")
    else:
        outliers_df = filtered[filtered["_outlier"]]
        if not outliers_df.empty:
            with st.expander(f"⚠️ {len(outliers_df)} Data Outlier Terdeteksi — klik untuk lihat detail"):
                show_cols_out = [c for c in ["Nomor", "Alamat", "Kecamatan", "Kota",
                                              "Tahun_Bersih", "Luas_Tanah", "Harga_Tanah"]
                                 if c in outliers_df.columns]
                st.dataframe(outliers_df[show_cols_out], use_container_width=True)

        all_cols = [c for c in filtered.columns if not c.startswith("_")]
        default_cols = [c for c in all_cols if c not in ("Latitude", "Longitude", "Tahun_Bersih")]
        disp_cols = st.multiselect("Pilih kolom yang ditampilkan:", all_cols, default=default_cols)

        if disp_cols:
            disp_df = filtered[disp_cols].copy()
            col_cfg = {}
            if "Harga_Tanah" in disp_cols:
                col_cfg["Harga_Tanah"] = st.column_config.NumberColumn(
                    "Harga Tanah (Rp/m²)", format="Rp %.0f"
                )
            if "Luas_Tanah" in disp_cols:
                col_cfg["Luas_Tanah"] = st.column_config.NumberColumn("Luas Tanah (m²)", format="%.0f m²")
            if "Luas_Bangunan" in disp_cols:
                col_cfg["Luas_Bangunan"] = st.column_config.NumberColumn("Luas Bangunan (m²)", format="%.0f m²")

            st.dataframe(disp_df, use_container_width=True, height=500, column_config=col_cfg)

            dl_col1, dl_col2 = st.columns([1, 5])
            with dl_col1:
                st.download_button(
                    label="📥 Download Excel",
                    data=to_excel_bytes(disp_df),
                    file_name=f"data_tanah_{city_label}_{selected_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ANALISA PERBANDINGAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_analisa:
    st.subheader("🔄 Analisa Perbandingan & Indikasi Nilai")
    st.markdown(
        "Pilih data pembanding, masukkan parameter koreksi, lalu sistem menghitung "
        "**harga indikasi** dan **koefisien variasi (CV)** sesuai standar penilaian."
    )

    if filtered.empty:
        st.warning("Tidak ada data yang sesuai dengan filter.")
    else:
        is_subject_mask = filtered["Nomor"].astype(str).str.lower().str.contains("obyek", na=False)
        subject_rows    = filtered[is_subject_mask]
        comparable_rows = filtered[~is_subject_mask]

        col_subj, col_comp = st.columns([1, 2])

        with col_subj:
            st.markdown("#### 🏠 Obyek Penilaian")
            if not subject_rows.empty:
                s = subject_rows.iloc[0]
                st.info(
                    f"**Alamat:** {s.get('Alamat', '-')}  \n"
                    f"**Kecamatan:** {s.get('Kecamatan', '-')}  \n"
                    f"**Luas Tanah:** {s.get('Luas_Tanah', '-')} m²  \n"
                    f"**Harga Indikasi Awal:** {format_currency(s.get('Harga_Tanah', 0))}/m²"
                )
                subj_luas  = float(s.get("Luas_Tanah") or 0)
                subj_harga = float(s.get("Harga_Tanah") or 0)
            else:
                st.warning("Tidak ada baris 'Obyek Penilaian' di data. Masukkan manual:")
                subj_luas  = st.number_input("Luas Tanah Obyek (m²)", value=0.0, min_value=0.0, step=10.0)
                subj_harga = st.number_input("Harga Indikasi Obyek (Rp/m²)", value=0.0, min_value=0.0, step=100000.0)

            st.markdown("#### ⚙️ Parameter Koreksi")
            ref_year      = st.number_input("Tahun Referensi Penilaian", value=2025, min_value=2000, max_value=2100, step=1)
            time_adj_pct  = st.number_input("Koreksi Waktu (%/tahun)", value=5.0, min_value=0.0, max_value=50.0, step=0.5,
                                             help="Kenaikan harga pasar per tahun (positif = pasar naik)")
            size_adj_pct  = st.number_input("Koreksi Luas (%/100m²)", value=2.0, min_value=0.0, max_value=20.0, step=0.5,
                                             help="Penyesuaian harga akibat perbedaan luas per 100m²")

        with col_comp:
            st.markdown("#### 📋 Pilih Data Pembanding")
            if comparable_rows.empty:
                st.warning("Tidak ada data pembanding tersedia.")
            else:
                nomor_opts = comparable_rows["Nomor"].astype(str).tolist()
                selected   = st.multiselect(
                    "Pilih nomor data pembanding (disarankan 3–5 data):",
                    options=nomor_opts,
                    default=nomor_opts[:min(3, len(nomor_opts))],
                )

                if selected:
                    comp = comparable_rows[comparable_rows["Nomor"].astype(str).isin(selected)].copy()

                    comp["Selisih_Tahun"]  = ref_year - comp["Tahun_Bersih"].fillna(ref_year)
                    comp["Kor_Waktu_%"]    = comp["Selisih_Tahun"] * time_adj_pct
                    comp["Harga_Kor_Waktu"] = comp["Harga_Tanah"] * (1 + comp["Kor_Waktu_%"] / 100)

                    if subj_luas > 0:
                        comp["Selisih_Luas"] = comp["Luas_Tanah"].fillna(subj_luas) - subj_luas
                        comp["Kor_Luas_%"]   = -(comp["Selisih_Luas"] / 100) * size_adj_pct
                    else:
                        comp["Kor_Luas_%"]   = 0.0

                    comp["Harga_Final"] = comp["Harga_Kor_Waktu"] * (1 + comp["Kor_Luas_%"] / 100)

                    # Summary table
                    tbl = comp[[
                        "Nomor", "Alamat", "Tahun_Bersih", "Luas_Tanah",
                        "Harga_Tanah", "Kor_Waktu_%", "Harga_Kor_Waktu",
                        "Kor_Luas_%", "Harga_Final"
                    ]].copy()
                    tbl.columns = [
                        "Nomor", "Alamat", "Tahun", "Luas (m²)",
                        "Harga Awal", "Kor. Waktu (%)", "Stl Kor. Waktu",
                        "Kor. Luas (%)", "Harga Final"
                    ]
                    st.dataframe(
                        tbl,
                        use_container_width=True,
                        column_config={
                            "Harga Awal":       st.column_config.NumberColumn(format="Rp %.0f"),
                            "Stl Kor. Waktu":   st.column_config.NumberColumn(format="Rp %.0f"),
                            "Harga Final":      st.column_config.NumberColumn(format="Rp %.0f"),
                            "Kor. Waktu (%)":   st.column_config.NumberColumn(format="%.1f %%"),
                            "Kor. Luas (%)":    st.column_config.NumberColumn(format="%.1f %%"),
                            "Luas (m²)":        st.column_config.NumberColumn(format="%.0f m²"),
                        },
                    )

                    # Results
                    harga_indikasi = comp["Harga_Final"].mean()
                    cv = (
                        comp["Harga_Final"].std() / comp["Harga_Final"].mean() * 100
                        if len(comp) > 1 and comp["Harga_Final"].mean() != 0
                        else 0.0
                    )

                    st.divider()
                    r1, r2, r3 = st.columns(3)
                    r1.metric("💰 Harga Indikasi",       format_currency(harga_indikasi) + "/m²")
                    r2.metric("📊 Koefisien Variasi (CV)", f"{cv:.1f}%")
                    r3.metric("📈 Jumlah Pembanding",     len(comp))

                    if cv <= 20:
                        st.success(f"✅ CV = {cv:.1f}% ≤ 20% → Homogen. Hasil dapat diandalkan.")
                    elif cv <= 30:
                        st.warning(f"⚠️ CV = {cv:.1f}% (20–30%) → Cukup beragam. Pertimbangkan tinjau ulang data.")
                    else:
                        st.error(f"❌ CV = {cv:.1f}% > 30% → Sangat beragam. Ganti/perbaiki data pembanding.")

                    # Bar chart comparison
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        x=comp["Nomor"].astype(str), y=comp["Harga_Tanah"],
                        name="Harga Awal", marker_color="#a8d8ea",
                    ))
                    fig.add_trace(go.Bar(
                        x=comp["Nomor"].astype(str), y=comp["Harga_Final"],
                        name="Harga Setelah Koreksi", marker_color="#667eea",
                    ))
                    if subj_harga > 0:
                        fig.add_hline(y=subj_harga, line_dash="dash", line_color="red",
                                      annotation_text=f"Harga Obyek: {format_currency(subj_harga)}")
                    fig.add_hline(y=harga_indikasi, line_dash="dot", line_color="green",
                                  annotation_text=f"Indikasi: {format_currency(harga_indikasi)}")
                    fig.update_layout(
                        title="Perbandingan Harga Data Pembanding",
                        barmode="group",
                        xaxis_title="Nomor Data",
                        yaxis_title="Harga (Rp/m²)",
                        margin=dict(t=50, b=20),
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # Download comparison result
                    st.download_button(
                        label="📥 Download Hasil Perbandingan",
                        data=to_excel_bytes(tbl),
                        file_name=f"analisa_perbandingan_{city_label}_{selected_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
