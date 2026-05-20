import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster, HeatMap
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
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

@st.cache_data(show_spinner="Memuat data...")
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    return df

df = load_data(file)
df["Tahun_Bersih"] = df["Tahun"].apply(bersihkan_tahun)

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

if city_input.strip():
    filtered = filtered[filtered["Kota"].astype(str).str.strip().str.lower().str.contains(
        city_input.strip().lower(), na=False
    )]
if selected_year != "Semua Tahun":
    filtered = filtered[filtered["Tahun_Bersih"] == int(selected_year)]
if selected_kecamatan != "Semua Kecamatan":
    filtered = filtered[filtered["Kecamatan"].astype(str) == selected_kecamatan]
if price_range:
    filtered = filtered[
        (filtered["Harga_Tanah"] >= price_range[0]) &
        (filtered["Harga_Tanah"] <= price_range[1])
    ]
if luas_range:
    filtered = filtered[
        (filtered["Luas_Tanah"] >= luas_range[0]) &
        (filtered["Luas_Tanah"] <= luas_range[1])
    ]

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
        # Sanity-check coordinate bounds
        map_df = map_df[
            map_df["Latitude"].between(-90, 90) &
            map_df["Longitude"].between(-180, 180)
        ]

        lat0 = map_df["Latitude"].mean() if not map_df.empty else -2.548926
        lon0 = map_df["Longitude"].mean() if not map_df.empty else 118.0148634

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

        if show_heatmap and not map_df.empty:
            heat = [
                [r.Latitude, r.Longitude, float(r.Harga_Tanah)]
                for r in map_df.itertuples()
                if pd.notna(getattr(r, "Harga_Tanah", None))
            ]
            if heat:
                HeatMap(heat, name="Heatmap Harga", radius=30, blur=20, min_opacity=0.4).add_to(m)

        if use_clustering:
            marker_layer = MarkerCluster(name="Data Pembanding").add_to(m)
        else:
            marker_layer = folium.FeatureGroup(name="Data Pembanding").add_to(m)

        for r in map_df.itertuples():
            nomor    = str(safe_get(r, "Nomor")).strip()
            tahun    = getattr(r, "Tahun_Bersih", 0) or 0
            is_subj  = "obyek" in nomor.lower()
            harga_fmt = format_currency(getattr(r, "Harga_Tanah", 0))

            if is_subj:
                folium.Marker(
                    location=[r.Latitude, r.Longitude],
                    tooltip="🏠 Obyek Penilaian — klik untuk detail",
                    icon=folium.DivIcon(
                        html="""
                        <div style="position:relative;width:38px;height:50px">
                          <div style="width:38px;height:38px;background:#c0392b;
                                      border-radius:50% 50% 50% 0;transform:rotate(-45deg);
                                      border:3px solid white;box-shadow:0 3px 8px rgba(0,0,0,0.5)">
                          </div>
                          <span style="position:absolute;top:4px;left:6px;
                                       font-size:18px;line-height:1">🏠</span>
                        </div>""",
                        icon_size=(38, 50), icon_anchor=(19, 50),
                    ),
                ).add_to(m)
                folium.Marker(
                    [r.Latitude, r.Longitude],
                    icon=folium.DivIcon(
                        html="""<div style="font-size:12px;font-weight:bold;color:#c0392b;
                                    background:rgba(255,255,255,0.95);padding:3px 8px;
                                    border-radius:5px;white-space:nowrap;border:2px solid #c0392b;
                                    box-shadow:1px 2px 5px rgba(0,0,0,0.3);pointer-events:none;
                                    margin-top:2px">🏠 OBYEK PENILAIAN</div>""",
                        icon_size=(175, 28), icon_anchor=(87, -4),
                    ),
                ).add_to(m)
            else:
                warna = get_color_by_year(tahun)
                folium.Marker(
                    location=[r.Latitude, r.Longitude],
                    tooltip=f"Data {nomor} | {harga_fmt}/m²",
                    icon=folium.Icon(color=warna, icon="info-sign", prefix="glyphicon"),
                ).add_to(marker_layer)
                label_color = "#922b21" if warna == "red" else "#154360"
                folium.Marker(
                    [r.Latitude, r.Longitude],
                    icon=folium.DivIcon(
                        html=f"""<div style="font-size:11px;color:{label_color};font-weight:bold;
                                    background:rgba(255,255,255,0.85);padding:2px 5px;
                                    border-radius:4px;white-space:nowrap;
                                    border:1px solid {label_color};pointer-events:none">
                            {harga_fmt}/m²<br>
                            <span style="font-size:10px">Data {nomor}</span>
                        </div>""",
                        icon_size=(140, 36), icon_anchor=(0, 0),
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
        </div>"""
        m.get_root().html.add_child(folium.Element(legend))

        # ── Split layout: peta kiri | panel detail kanan ────────────────
        col_map, col_detail = st.columns([3, 2], gap="medium")

        with col_map:
            n_subj = map_df["Nomor"].astype(str).str.lower().str.contains("obyek").sum()
            n_comp = len(map_df) - n_subj
            st.caption(
                f"**{n_subj}** Obyek Penilaian + **{n_comp}** Data Pembanding"
                f" — klik marker untuk detail di panel kanan"
            )
            result = st_folium(
                m, width="100%", height=620,
                returned_objects=["last_object_clicked"],
                key="folium_peta",
            )
            # Simpan marker yang diklik ke session_state
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

        # ── Panel detail ─────────────────────────────────────────────────
        with col_detail:
            sel_idx = st.session_state.get("peta_sel")
            if sel_idx is None or sel_idx not in map_df.index:
                st.markdown("""
                <div style="text-align:center;padding:60px 20px;color:#bbb;
                            border:2px dashed #e0e0e0;border-radius:10px;margin-top:40px">
                  <div style="font-size:48px">👆</div>
                  <div style="font-size:14px;margin-top:12px;line-height:1.6">
                    Klik marker di peta<br>untuk melihat detail properti
                  </div>
                </div>""", unsafe_allow_html=True)
            else:
                row      = map_df.loc[sel_idx]
                nomor_d  = str(safe_get(row, "Nomor")).strip()
                is_s     = "obyek" in nomor_d.lower()
                tahun_d  = row.get("Tahun_Bersih") if hasattr(row, "get") else getattr(row, "Tahun_Bersih", None)
                harga_d  = format_currency(getattr(row, "Harga_Tanah", None))
                foto_url = str(safe_get(row, "Foto", ""))
                lh3_url, thumb_url = gdrive_thumbnail(foto_url, width=500)

                # Tombol tutup
                if st.button("✕ Tutup", key="tutup_detail"):
                    st.session_state.pop("peta_sel", None)
                    st.rerun()

                # Header
                if is_s:
                    st.markdown(
                        '<div style="background:#fdecea;border-left:4px solid #c0392b;'
                        'padding:8px 12px;border-radius:0 6px 6px 0;margin-bottom:8px">'
                        '<b style="color:#c0392b;font-size:15px">🏠 Obyek Penilaian</b></div>',
                        unsafe_allow_html=True)
                else:
                    jenis_d = safe_get(row, "Jenis_Data")
                    badge   = "#e67e22" if "penawaran" in str(jenis_d).lower() else "#2980b9"
                    tahun_lbl = f"· {int(tahun_d)}" if tahun_d and not pd.isna(tahun_d) else ""
                    st.markdown(
                        f'<div style="background:#eafaf1;border-left:4px solid #27ae60;'
                        f'padding:8px 12px;border-radius:0 6px 6px 0;margin-bottom:8px;'
                        f'display:flex;align-items:center;gap:8px">'
                        f'<b style="color:#1a7a4a;font-size:15px">Data {nomor_d}</b>'
                        f'<span style="background:{badge};color:white;border-radius:10px;'
                        f'padding:1px 8px;font-size:11px">{jenis_d}</span>'
                        f'<span style="color:#999;font-size:12px;margin-left:auto">{tahun_lbl}</span>'
                        f'</div>',
                        unsafe_allow_html=True)

                # Foto
                if lh3_url:
                    st.markdown(
                        f'<img src="{lh3_url}" '
                        f'style="width:100%;border-radius:8px;border:1px solid #e0e0e0;'
                        f'box-shadow:0 2px 8px rgba(0,0,0,0.12);margin-bottom:4px" '
                        f'referrerpolicy="no-referrer" '
                        f'onerror="this.src=\'{thumb_url}\';this.onerror=null;">',
                        unsafe_allow_html=True)

                # Galeri foto obyek (foto mini horizontal)
                if is_s:
                    foto_extra = [
                        (safe_get(row, "Foto_Dalam",         ""), "Dalam"),
                        (safe_get(row, "Foto_Samping_Kanan", ""), "Kanan"),
                        (safe_get(row, "Foto_Samping_Kiri",  ""), "Kiri"),
                        (safe_get(row, "Gambar_Situasi",     ""), "Situasi"),
                    ]
                    valid_extra = [(u, lb) for u, lb in foto_extra
                                   if u not in ("", "#", "-", "nan", "None")]
                    if valid_extra:
                        gcols = st.columns(len(valid_extra))
                        for gc, (u, lb) in zip(gcols, valid_extra):
                            l3, th = gdrive_thumbnail(u, width=200)
                            if l3:
                                gc.markdown(
                                    f'<img src="{l3}" style="width:100%;border-radius:5px;'
                                    f'border:1px solid #ddd" referrerpolicy="no-referrer" '
                                    f'onerror="this.src=\'{th}\';this.onerror=null;">',
                                    unsafe_allow_html=True)
                                gc.caption(lb)

                # Link bar
                links_html = []
                if lh3_url and foto_url not in ("", "#", "-", "nan", "None"):
                    links_html.append(f'<a href="{foto_url}" target="_blank" '
                                      f'style="color:#2980b9;text-decoration:none;font-size:12px">'
                                      f'&#128247; Foto Depan &#8599;</a>')
                fj = str(safe_get(row, "Foto_Jalan", ""))
                if fj not in ("", "#", "-", "nan", "None"):
                    links_html.append(f'<a href="{fj}" target="_blank" '
                                      f'style="color:#2980b9;text-decoration:none;font-size:12px">'
                                      f'&#128247; Foto Jalan &#8599;</a>')
                sv_url_d = generate_streetview_url(
                    safe_get(row, "Latitude"), safe_get(row, "Longitude"))
                links_html.append(f'<a href="{sv_url_d}" target="_blank" '
                                   f'style="color:#2980b9;text-decoration:none;font-size:12px">'
                                   f'&#128269; Street View &#8599;</a>')
                st.markdown(
                    '<div style="display:flex;gap:14px;margin:6px 0 10px;flex-wrap:wrap">'
                    + " ".join(links_html) + "</div>",
                    unsafe_allow_html=True)

                # Kartu harga (Data Pembanding)
                if not is_s:
                    ht = getattr(row, "Harga_Total", None)
                    ht_str = format_currency(ht) if (ht and not pd.isna(ht)) else None
                    c1, c2 = st.columns(2)
                    if ht_str:
                        c1.metric("Harga Total", ht_str)
                    c2.metric("Harga/m²", harga_d)

                # ── Info sections ─────────────────────────────────────────
                def info_row(label, val):
                    v = str(val) if val not in ("-", None, "") else "—"
                    st.markdown(
                        f'<div style="display:flex;border-bottom:1px solid #f0f0f0;'
                        f'padding:4px 0;font-size:12.5px">'
                        f'<span style="color:#888;min-width:110px;flex-shrink:0">{label}</span>'
                        f'<span style="color:#222">{v}</span></div>',
                        unsafe_allow_html=True)

                def info_2col(l1, v1, l2, v2):
                    c1, c2 = st.columns(2)
                    with c1:
                        info_row(l1, v1)
                    with c2:
                        info_row(l2, v2)

                def sec_hdr(icon, title):
                    st.markdown(
                        f'<div style="background:#f0faf5;color:#1a7a4a;font-size:10px;'
                        f'font-weight:bold;text-transform:uppercase;letter-spacing:0.6px;'
                        f'padding:3px 8px;border-radius:3px;margin:10px 0 4px">'
                        f'{icon} {title}</div>',
                        unsafe_allow_html=True)

                if is_s:
                    sec_hdr("📋", "Identitas")
                    info_2col("Pemilik",       safe_get(row,"Pemilik"),
                              "Jenis",          safe_get(row,"Jenis_Properti"))
                    info_2col("Kode Inspeksi", safe_get(row,"Kode_Inspeksi"),
                              "Reviewer",       safe_get(row,"Reviewer"))
                    info_row("Pemberi Tugas",  safe_get(row,"Pemberi_Tugas"))
                else:
                    sec_hdr("📋", "Properti")
                    info_2col("Jenis",     safe_get(row,"Jenis_Properti"),
                              "Tahun",     int(tahun_d) if tahun_d and not pd.isna(tahun_d) else "—")

                sec_hdr("📍", "Lokasi")
                info_row("Alamat",    safe_get(row,"Alamat"))
                info_row("Kompleks",  safe_get(row,"Kompleks"))
                info_2col("Kelurahan", safe_get(row,"Kelurahan"),
                          "Kecamatan", safe_get(row,"Kecamatan"))
                info_2col("Kota",      safe_get(row,"Kota"),
                          "Propinsi",  safe_get(row,"Propinsi"))

                sec_hdr("📐", "Fisik")
                def _luas(col):
                    v = getattr(row, col, None)
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return "—"
                    try:
                        return f"{float(v):,.0f} m²".replace(",", ".")
                    except Exception:
                        return str(v)
                info_2col("Luas Tanah",   _luas("Luas_Tanah"),
                          "Luas Bangunan", _luas("Luas_Bangunan"))
                if not is_s:
                    info_2col("Kondisi Bgn", safe_get(row,"Kondisi_Bangunan"),
                              "Kelas Bgn",   safe_get(row,"Kelas_Bangunan"))
                info_2col("Peruntukan",  safe_get(row,"Peruntukan"),
                          "Kepemilikan", safe_get(row,"Kepemilikan"))

                if not is_s:
                    sec_hdr("📞", "Kontak")
                    info_2col("Nama",  safe_get(row,"Kontak"),
                              "Telp",  safe_get(row,"Telp"))

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
