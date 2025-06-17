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

required = {"Kota","Harga_Tanah","Latitude","Longitude"}
if not required.issubset(df.columns):
    st.error(f"❌ File harus punya kolom: {', '.join(required)}")
    st.stop()

city = st.text_input("🔍 Masukkan nama kota untuk mencari data:")
filtered = df[df["Kota"].str.strip().str.lower() == city.strip().lower()] if city else df

st.dataframe(filtered, use_container_width=True)

st.subheader("📌 Peta Lokasi Properti")

# Center map
if not filtered.empty:
    lat0, lon0 = filtered["Latitude"].mean(), filtered["Longitude"].mean()
else:
    lat0, lon0 = -2.548926, 118.0148634

m = folium.Map(location=[lat0, lon0], zoom_start=10 if city else 5)


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
        warna = get_color_by_year(tahun)
        
        tooltip = (
            f"{r.Nomor_Data}<br>"
            f"{r.Kelurahan}<br>"
            f"{r.Kecamatan}<br>"
            f"{r.Kota}<br>"
            f"Harga Tanah: <b>{format_currency(r.Harga_Tanah)}</b>/m²"
        )
        # 1. Marker biasa
        folium.Marker(
            location=[r.Latitude, r.Longitude],
            tooltip=tooltip,
            icon=folium.Icon(color=warna)
        ).add_to(m)

        # 2. Label tetap tampil (DivIcon)
        folium.map.Marker(
            [r.Latitude, r.Longitude],
            icon=folium.DivIcon(
                html=f"""
                <div style='font-size:12px;
                            color:{warna};
                            font-weight:bold;
                            background-color:None;
                            padding:2px 4px;
                            border-radius:4px;
                            white-space: nowrap;'>
                    {f"<b>{format_currency(r.Harga_Tanah)}</b>/m²"}
                </div>
                """
            )
        ).add_to(m)

st_folium(m, width=1500, height=700)
