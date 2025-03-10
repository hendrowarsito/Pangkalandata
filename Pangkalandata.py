import pandas as pd
import streamlit as st
import folium
from streamlit_folium import st_folium

# **Mengatur tampilan agar memenuhi layar**
st.set_page_config(layout="wide")

st.title("📍 Pangkalan Data Tanah KJPP Suwendho Rinaldy dan Rekan 🏡")

# Fungsi untuk memformat angka ke mata uang Rupiah
def format_currency(value):
    return f"Rp {value:,.0f}".replace(',', '.')

# Upload file Excel
file = st.file_uploader("📂 Unggah file Excel berisi data tanah", type=["xlsx"])

if file is not None:
    # Membaca file Excel yang diunggah
    df = pd.read_excel(file)

    # Validasi apakah file memiliki kolom yang diperlukan
    required_columns = {"Kota", "Harga_Tanah", "Latitude", "Longitude"}
    if not required_columns.issubset(df.columns):
        st.error("❌ File Excel harus memiliki kolom: Kota, Harga_Tanah, Latitude, Longitude")
        st.stop()

    # Input untuk pencarian kota
    name_input = st.text_input("🔍 Masukkan nama kota untuk mencari data:")

    # Filter data berdasarkan input pengguna
    filtered_data = df[df['Kota'].str.contains(name_input, case=False, na=False)] if name_input else df

    st.dataframe(filtered_data, use_container_width=True)

    # **Menampilkan peta dan tabel properti secara lebar**
    col1, col2 = st.columns([3, 2])  # Peta lebih besar (3 bagian) | Tabel properti lebih kecil (2 bagian)

    with col1:
        st.subheader("📌 Peta Lokasi Properti")

        # **Menentukan pusat peta berdasarkan kota yang dicari**
        if not filtered_data.empty:
            lat_center = filtered_data["Latitude"].mean()
            lon_center = filtered_data["Longitude"].mean()
        else:
            lat_center, lon_center = -2.548926, 118.0148634  # Default Indonesia jika tidak ditemukan

        # **Cegah error jika lat_center atau lon_center NaN**
        if pd.isna(lat_center) or pd.isna(lon_center):
            lat_center, lon_center = -2.548926, 118.0148634  # Default Indonesia

        # Inisialisasi peta dengan pusat berdasarkan pencarian kota
        m = folium.Map(location=[lat_center, lon_center], zoom_start=10 if name_input else 5)

        # Menambahkan marker berdasarkan data yang diunggah
        for index, row in df.iterrows():
            if not pd.isna(row['Latitude']) and not pd.isna(row['Longitude']):
                formatted_price = format_currency(row['Harga_Tanah'])
                folium.Marker(
                    location=[row['Latitude'], row['Longitude']],
                    popup=f"<b>{row['Kota']}</b><br>Harga Tanah: <b>{formatted_price}</b>/m²",
                    tooltip=row['Kota']
                ).add_to(m)

        # Menampilkan peta dalam Streamlit
        map_data = st_folium(m, width=1200, height=600)

    with col2:
        st.subheader("📋 Data Properti di Peta")

        # Menampilkan data properti dalam batas tampilan peta
        if map_data and "bounds" in map_data and map_data["bounds"]:
            bounds = map_data["bounds"]
            if "_southWest" in bounds and "_northEast" in bounds:
                lat_min = bounds["_southWest"]["lat"]
                lat_max = bounds["_northEast"]["lat"]
                lon_min = bounds["_southWest"]["lng"]
                lon_max = bounds["_northEast"]["lng"]

                # Filter properti dalam tampilan peta
                visible_data = df[
                    (df["Latitude"] >= lat_min) & (df["Latitude"] <= lat_max) &
                    (df["Longitude"] >= lon_min) & (df["Longitude"] <= lon_max)
                ].copy()

                if not visible_data.empty:
                    visible_data["Harga Tanah (Rp/m²)"] = visible_data["Harga_Tanah"].apply(format_currency)
                    st.dataframe(visible_data[["Kota", "Harga Tanah (Rp/m²)", "Latitude", "Longitude"]], use_container_width=True)
                else:
                    st.write("🔍 Tidak ada properti yang terlihat pada peta.")
            else:
                st.write("⚠ Struktur bounds tidak sesuai. Silakan zoom atau geser peta.")
        else:
            st.write("📍 Pindahkan atau zoom peta untuk menampilkan data properti.")
