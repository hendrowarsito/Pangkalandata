import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
import os

st.set_page_config(layout="wide", initial_sidebar_state="expanded")

# Sidebar input
st.sidebar.markdown("## Pangkalan Data Tanah KJPP Suwendho Rinaldy dan Rekan 🏡")
st.sidebar.header("🔧 Filter Data")
file = st.sidebar.file_uploader("📂 Unggah file Excel atau CSV data tanah", type=["xlsx", "csv"])

# Reset tombol jika file baru diunggah
if file and "last_file" in st.session_state and file != st.session_state["last_file"]:
    st.session_state["tampilkan"] = False
st.session_state["last_file"] = file

if not file:
    st.sidebar.info("Silakan unggah file Excel atau CSV untuk melanjutkan.")
    st.stop()

@st.cache_data(show_spinner=False)
def load_data(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(uploaded_file)
    elif ext in [".xls", ".xlsx"]:
        df = pd.read_excel(uploaded_file)
    else:
        st.error("Format file tidak didukung.")
        return pd.DataFrame()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    return df

def format_currency(value):
    return f"Rp {value:,.0f}".replace(",", ".")

def bersihkan_tahun(val):
    try:
        val = str(val).replace(",", "").strip()
        return pd.to_numeric(val, errors="coerce", downcast="integer")
    except:
        return None

# Load dan bersihkan data
df = load_data(file)
if df.empty:
    st.stop()

ext = os.path.splitext(file.name)[1].lower()
df["Tahun_Bersih"] = df["Tahun"].apply(bersihkan_tahun) if "Tahun" in df.columns else None

city_input = st.sidebar.text_input("🔍 Masukkan nama kota:")
available_years = sorted([int(y) for y in df["Tahun_Bersih"].dropna().unique()], reverse=True) if "Tahun_Bersih" in df.columns else []
tahun_opsi = ["Semua Tahun"] + [str(t) for t in available_years]
selected_year = st.sidebar.selectbox("📅 Pilih Tahun Data:", tahun_opsi)

# Tombol dan session state
if "tampilkan" not in st.session_state:
    st.session_state["tampilkan"] = False

if st.sidebar.button("Tampilkan Data"):
    st.session_state["tampilkan"] = True

if not st.session_state["tampilkan"]:
    st.sidebar.info("Tekan tombol 'Tampilkan Data' untuk melihat hasil.")
    st.stop()

city_clean = city_input.strip().lower()
filtered = df.copy()
if "Kota" in df.columns and city_input:
    filtered = filtered[filtered["Kota"].str.strip().str.lower() == city_clean]
if "Tahun_Bersih" in df.columns and selected_year != "Semua Tahun":
    filtered = filtered[filtered["Tahun_Bersih"] == int(selected_year)]

st.success(f"Menampilkan {len(filtered)} data untuk kota '{city_input}' dan tahun '{selected_year}'")

# Tabs
tabs = st.tabs(["🗺️ Peta XLSX", "📋 Tabel XLSX", "🗺️ Peta CSV", "📋 Tabel CSV"])

# Fungsi warna marker

def get_color_by_year(year):
    if pd.isna(year):
        return "gray"
    if year >= 2025:
        return "green"
    elif year >= 2024:
        return "blue"
    elif year >= 2023:
        return "orange"
    else:
        return "red"

# Fungsi tampilkan peta

def tampilkan_peta(data):
    if not data.empty:
        lat0, lon0 = data["Latitude"].mean(), data["Longitude"].mean()
    else:
        lat0, lon0 = -2.548926, 118.0148634

    m = folium.Map(
        location=[lat0, lon0],
        zoom_start=8,
        min_zoom=5,
        max_zoom=18,
        prefer_canvas=True,
        scrollWheelZoom=True,
        control_scale=True
    )

    folium.TileLayer('OpenStreetMap', name='OpenStreetMap', control=True).add_to(m)
    folium.TileLayer(
        tiles='https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
        name='Voyager Terrain',
        attr='©OpenStreetMap contributors ©CartoDB',
        overlay=False,
        control=True
    ).add_to(m)
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
        name='Satellite',
        attr='Tiles © Esri',
        overlay=False,
        control=True
    ).add_to(m)

    folium.LayerControl(collapsed=False).add_to(m)

    for r in data.itertuples():
        if pd.notna(r.Latitude) and pd.notna(r.Longitude):
            tahun = getattr(r, "Tahun", None)
            nomor = str(getattr(r, "Nomor", "")).strip()
            warna = get_color_by_year(tahun)
            warna_teks = "red" if nomor.lower() == "obyek penilaian" else warna
            foto_link = getattr(r, "Foto", "#") or "#"
            popup = f"<b>{getattr(r, 'Kontak', '')}</b><br><b>{getattr(r, 'Telp', '')}</b><br>"
            tooltip = f"{nomor}<br>Tahun: {tahun}<br>Alamat: {getattr(r, 'Alamat', '')}<br>Kota: {getattr(r, 'Kota', '')}"
            folium.Marker(
                location=[r.Latitude, r.Longitude],
                popup=popup,
                tooltip=tooltip,
                icon=folium.Icon(color=warna)
            ).add_to(m)

            folium.map.Marker(
                [r.Latitude, r.Longitude],
                icon=folium.DivIcon(
                    html=f"""
                    <div style='font-size:12px;
                                color:{warna_teks};
                                font-weight:bold;
                                background-color:transparent;
                                padding:2px 4px;
                                border-radius:4px;
                                white-space: nowrap;'>
                        {format_currency(getattr(r, 'Harga_Tanah', 0))}/m²<br>
                        <a href=\"{foto_link}\" target=\"_blank\" style=\"color:{warna_teks}; text-decoration:underline;\">
                            {nomor}
                        </a>
                    </div>
                    """
                )
            ).add_to(m)

    st_folium(m, width=1300, height=700)

# Tampilkan tab
with tabs[0]:
    if ext != ".csv":
        tampilkan_peta(filtered)
with tabs[1]:
    if ext != ".csv":
        st.dataframe(filtered, use_container_width=True)
with tabs[2]:
    if ext == ".csv":
        tampilkan_peta(filtered)
with tabs[3]:
    if ext == ".csv":
        st.dataframe(filtered, use_container_width=True)
