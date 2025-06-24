import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

st.set_page_config(layout="wide")
st.title("📍 Pangkalan Data Tanah KJPP Suwendho Rinaldy dan Rekan 🏡")

# cache the Excel load so it only runs once per file upload
@st.cache_data(show_spinner=False)
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    # ensure numeric lat/lon
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    return df

def format_currency(value):
    return f"Rp {value:,.0f}".replace(",", ".")

file = st.file_uploader("📂 Unggah file Excel berisi data tanah", type=["xlsx"])
if not file:
    st.info("Silakan unggah file Excel untuk melanjutkan.")
    st.stop()

df = load_data(file)


# Bersihkan kolom Tahun agar jadi integer
def bersihkan_tahun(val):
    try:
        val = str(val).replace(",", "").strip()
        return pd.to_numeric(val, errors="coerce", downcast="integer")
    except:
        return None

df["Tahun_Bersih"] = df["Tahun"].apply(bersihkan_tahun)

# --- Input nama kota dan dropdown tahun ---
city = st.text_input("🔍 Masukkan nama kota untuk mencari data:")

# Ambil semua tahun valid dari Tahun_Bersih
available_years = sorted([int(y) for y in df["Tahun_Bersih"].dropna().unique()], reverse=True)
tahun_opsi = ["Semua Tahun"] + [str(t) for t in available_years]

selected_year = st.selectbox("📅 Pilih Tahun Data:", tahun_opsi)

# --- Filter DataFrame berdasarkan input kota dan tahun ---
city_clean = city.strip().lower()
filtered = df.copy()

if city:
    filtered = filtered[filtered["Kota"].str.strip().str.lower() == city_clean]
if selected_year != "Semua Tahun":
    filtered = filtered[filtered["Tahun_Bersih"] == int(selected_year)]

st.success(f"Menampilkan {len(filtered)} data untuk kota '{city}' dan tahun '{selected_year}'")
st.dataframe(filtered)    
#st.dataframe(filtered, use_container_width=True)

st.subheader("📌 Peta Lokasi Properti")

# Center map
if not filtered.empty:
    lat0, lon0 = filtered["Latitude"].mean(), filtered["Longitude"].mean()
else:
    lat0, lon0 = -2.548926, 118.0148634

m = folium.Map(
    location=[lat0, lon0],
    zoom_start=5 if city else 5,
    min_zoom=1,
    max_zoom=18,
    prefer_canvas=True,
    scrollWheelZoom=True,
    control_scale=True
)

# Tambahkan beberapa layer
folium.TileLayer('OpenStreetMap', name='OpenStreetMap').add_to(m)

# Terrain Layer (Stamen Terrain)
folium.TileLayer(
    tiles='https://stamen-tiles.a.ssl.fastly.net/terrain/{z}/{x}/{y}.jpg',
    attr='Map tiles by Stamen Design, under CC BY 3.0. Data by OpenStreetMap, under ODbL.',
    name='Terrain',
    overlay=False,
    control=True
).add_to(m)

# Satellite Layer (Google Satellite via ESRI World Imagery)
folium.TileLayer(
    tiles='https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    attr='Tiles © Esri — Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community',
    name='Satellite',
    overlay=False,
    control=True
).add_to(m)

# Tambahkan kontrol pemilih layer
folium.LayerControl(collapsed=False).add_to(m)

# Warna berdasarkan tahun
def get_color_by_year(year):
    if year >= 2025:
        return "green"
    elif year >= 2024:
        return "blue"
    elif year >= 2023:
        return "orange"
    else:
        return "red"

# Tambahkan marker langsung (tanpa cluster)
for r in filtered.itertuples():
    if pd.notna(r.Latitude) and pd.notna(r.Longitude):
        tahun = getattr(r, "Tahun", 0)  # asumsi kolom 'Tahun' ada
        # Penyesuaian warna dan link
        nomor = str(getattr(r, "Nomor", "")).strip()
        warna = get_color_by_year(tahun)
        warna_teks = "red" if nomor.lower() == "obyek penilaian" else warna
        foto_link = getattr(r, "Foto", "#") or "#"
        popup = (
            f"<b>{r.Kontak}</b><br>"
            f"<b>{r.Telp}</b><br>"
            
        )
        tooltip = (
            f"{r.Nomor}</b><br>"
            f"Tahun: {tahun}<br>"
            f"Alamat: {r.Alamat}</b><br>"
            f"Kelurahan: {r.Kelurahan}<br>"
            f"Kecamatan: {r.Kecamatan}<br>"
            f"Kota: {r.Kota}<br>"
            f"Luas: Tanah {r.Luas_Tanah}</b> m²<br>"
            f"Luas Bangunan: {r.Luas_Bangunan}</b> m²<br>"
            f"Harga Tanah: <b>{format_currency(r.Harga_Tanah)}</b>/m²"
        )
        # 1. Marker biasa
        folium.Marker(
            location=[r.Latitude, r.Longitude],
            popup=popup,
            tooltip=tooltip,
            icon=folium.Icon(color=warna)
        ).add_to(m)

        # 2. Label tetap tampil (DivIcon)
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
                    {format_currency(r.Harga_Tanah)}/m²
                    <br><a href="{foto_link}" target="_blank" style="color:{warna_teks}; text-decoration:underline;">
                        {nomor}
                    </a>
                </div>
                """
            )
        ).add_to(m)

st_folium(m, width=1600, height=700)

