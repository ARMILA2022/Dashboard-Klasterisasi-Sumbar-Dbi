import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import geopandas as gpd
import zipfile
import tempfile
import os
import re
import unicodedata
import json
from pathlib import Path

# ============================================================
# 1. KONFIGURASI DASHBOARD
# ============================================================

st.set_page_config(
    page_title="Dashboard Klasterisasi Sumatera Barat",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_ZIP_PETA = "Provinsi Sumatera Barat-KAB_KOTA.zip"

# Warna hasil klaster penelitian dipertahankan
WARNA_KLASTER = {
    "1": "#36A269",
    "2": "#F2C94C",
    "3": "#E45B5B"
}

# Warna utama tampilan dashboard
BIRU_UTAMA = "#174A7E"
BIRU_TUA = "#102F50"
BIRU_MUDA = "#EAF3FB"
BIRU_AKSI = "#2878B5"
ABU_TEKS = "#536477"

# ============================================================
# 2. DESAIN DASHBOARD
# ============================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #F4F7FB;
        color: #1E293B;
    }

    [data-testid="stHeader"] {
        background-color: rgba(244, 247, 251, 0.95);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #102F50 0%, #174A7E 100%);
    }

    [data-testid="stSidebar"] * {
        color: #FFFFFF;
    }

    [data-testid="stSidebar"] [data-baseweb="select"] * {
        color: #1E293B;
    }

    [data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #2878B5;
    }

    .dashboard-header {
        background: linear-gradient(135deg, #102F50 0%, #2878B5 100%);
        color: white;
        padding: 30px;
        border-radius: 18px;
        min-height: 235px;
        box-shadow: 0 8px 24px rgba(16, 47, 80, 0.15);
    }

    .dashboard-header .eyebrow {
        font-size: 12px;
        letter-spacing: 2px;
        font-weight: 700;
        color: #D6E9FA;
        margin-bottom: 14px;
    }

    .dashboard-header h1 {
        color: white;
        font-size: 31px;
        line-height: 1.25;
        margin: 0 0 15px 0;
        font-weight: 800;
    }

    .dashboard-header p {
        color: #EDF6FF;
        font-size: 15px;
        line-height: 1.7;
        margin-bottom: 0;
    }

    .section-heading {
        color: #174A7E;
        font-size: 23px;
        font-weight: 800;
        margin-top: 20px;
        margin-bottom: 5px;
    }

    .section-subtitle {
        color: #64748B;
        font-size: 14px;
        margin-bottom: 18px;
    }

    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E0EAF4;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 4px 14px rgba(16, 47, 80, 0.05);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748B;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #174A7E;
        font-weight: 800;
    }

    div[data-testid="stTabs"] button {
        font-weight: 650;
    }

    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #174A7E;
    }

    .info-box {
        background: #EAF3FB;
        border-left: 5px solid #2878B5;
        padding: 15px 18px;
        border-radius: 8px;
        color: #23415F;
        margin: 10px 0 18px 0;
        line-height: 1.7;
    }

    .footer-box {
        text-align: center;
        color: #64748B;
        font-size: 12px;
        padding: 20px 10px;
        margin-top: 30px;
        border-top: 1px solid #DCE6F0;
    }

    .stDownloadButton button {
        background-color: #174A7E;
        color: white;
        border-radius: 8px;
        border: none;
    }

    .stDownloadButton button:hover {
        background-color: #2878B5;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# 3. FUNGSI BANTU
# ============================================================

def normalisasi_nama(nama):
    """Menyamakan format nama wilayah agar mudah dicocokkan."""
    if pd.isna(nama):
        return ""

    nama = str(nama).strip().lower()
    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter for karakter in nama
        if not unicodedata.combining(karakter)
    )

    nama = nama.replace("kabupaten", "")
    nama = nama.replace("kota administrasi", "")
    nama = nama.replace("kota", "")
    nama = nama.replace("kab.", "")
    nama = nama.replace("kab ", "")
    nama = nama.replace("kepulauan", "kepulauan")

    nama = re.sub(r"[^a-z0-9\s]", " ", nama)
    nama = re.sub(r"\s+", " ", nama).strip()

    # Penyesuaian nama wilayah yang sering berbeda penulisannya
    nama = nama.replace("sawah lunto", "sawahlunto")
    nama = nama.replace("padang panjang", "padang panjang")
    nama = nama.replace("pariaman", "pariaman")

    return nama


def cari_kolom(data, daftar_nama):
    """Mencari kolom berdasarkan beberapa kemungkinan nama."""
    kolom_normal = {
        str(k).strip().lower(): k
        for k in data.columns
    }

    for nama in daftar_nama:
        if nama.lower() in kolom_normal:
            return kolom_normal[nama.lower()]

    for kolom in data.columns:
        kolom_bersih = str(kolom).strip().lower()
        for nama in daftar_nama:
            if nama.lower() in kolom_bersih:
                return kolom

    return None


def siapkan_kolom_wilayah(data):
    """Memastikan tabel mempunyai kolom Kabupaten/Kota."""
    data = data.copy()

    kolom_wilayah = cari_kolom(
        data,
        ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota", "Wilayah", "Daerah"]
    )

    if kolom_wilayah is None:
        if data.index.name is not None:
            data = data.reset_index()
        elif not isinstance(data.index, pd.RangeIndex):
            data = data.reset_index()

        kolom_wilayah = cari_kolom(
            data,
            ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota", "Wilayah", "Daerah", "index"]
        )

    if kolom_wilayah is not None and kolom_wilayah != "Kabupaten/Kota":
        data = data.rename(columns={kolom_wilayah: "Kabupaten/Kota"})

    return data


def siapkan_kolom_klaster(data):
    """Menyamakan nama kolom klaster menjadi Klaster."""
    data = data.copy()

    kolom_klaster = cari_kolom(
        data,
        ["Klaster", "Cluster", "Label Klaster", "Hasil Klaster"]
    )

    if kolom_klaster is not None and kolom_klaster != "Klaster":
        data = data.rename(columns={kolom_klaster: "Klaster"})

    if "Klaster" in data.columns:
        data["Klaster"] = (
            data["Klaster"]
            .astype(str)
            .str.extract(r"(\d+)", expand=False)
        )

    return data


def format_angka(nilai, desimal=3):
    """Mengubah angka menjadi format yang mudah dibaca."""
    try:
        return f"{float(nilai):,.{desimal}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "-"


# ============================================================
# 4. MEMBACA DATA EXCEL
# ============================================================

if not Path(FILE_EXCEL).exists():
    st.error(
        f"File Excel tidak ditemukan: {FILE_EXCEL}. "
        "Pastikan file Excel sudah diunggah ke repositori GitHub "
        "yang sama dengan app.py."
    )
    st.stop()

try:
    excel = pd.ExcelFile(FILE_EXCEL)
    daftar_sheet = excel.sheet_names

    def baca_sheet(nama_sheet, **kwargs):
        if nama_sheet in daftar_sheet:
            return pd.read_excel(FILE_EXCEL, sheet_name=nama_sheet, **kwargs)
        return pd.DataFrame()

    data_klaster = baca_sheet("Hasil_Klaster")
    data_dbi = baca_sheet("Evaluasi_DBI")
    data_profil = baca_sheet("Profil_Klaster", index_col=0)
    data_pca = baca_sheet("Koordinat_PCA")

except Exception as kesalahan:
    st.error(f"Terjadi kesalahan saat membaca file Excel: {kesalahan}")
    st.stop()

if data_klaster.empty:
    st.error(
        "Sheet 'Hasil_Klaster' tidak ditemukan atau tidak berisi data. "
        "Periksa nama sheet pada file Excel."
    )
    st.stop()

data_klaster = siapkan_kolom_wilayah(data_klaster)
data_klaster = siapkan_kolom_klaster(data_klaster)

if "Kabupaten/Kota" not in data_klaster.columns:
    st.error(
        "Kolom nama wilayah tidak ditemukan pada sheet Hasil_Klaster. "
        "Pastikan ada kolom Kabupaten/Kota."
    )
    st.stop()

if "Klaster" not in data_klaster.columns:
    st.error(
        "Kolom Klaster tidak ditemukan pada sheet Hasil_Klaster. "
        "Pastikan hasil pengelompokan memiliki kolom Klaster atau Cluster."
    )
    st.stop()

data_klaster = data_klaster.dropna(subset=["Kabupaten/Kota", "Klaster"]).copy()
data_klaster["Kabupaten/Kota"] = data_klaster["Kabupaten/Kota"].astype(str).str.strip()
data_klaster["Klaster"] = data_klaster["Klaster"].astype(str)

# Menyamakan format kolom untuk data lain
if not data_dbi.empty:
    data_dbi.columns = [str(k).strip() for k in data_dbi.columns]

if not data_profil.empty:
    data_profil = data_profil.reset_index()
    data_profil = siapkan_kolom_klaster(data_profil)

if not data_pca.empty:
    data_pca = siapkan_kolom_wilayah(data_pca)
    data_pca = siapkan_kolom_klaster(data_pca)

# ============================================================
# 5. MENENTUKAN HASIL EVALUASI DBI
# ============================================================

kolom_k = cari_kolom(data_dbi, ["K", "Jumlah Klaster", "Jumlah Cluster"])
kolom_nilai_dbi = cari_kolom(data_dbi, ["DBI", "Davies-Bouldin Index", "Davies Bouldin Index"])

k_optimal = data_klaster["Klaster"].nunique()
dbi_terbaik = None

if not data_dbi.empty and kolom_k is not None and kolom_nilai_dbi is not None:
    data_dbi[kolom_k] = pd.to_numeric(data_dbi[kolom_k], errors="coerce")
    data_dbi[kolom_nilai_dbi] = pd.to_numeric(
        data_dbi[kolom_nilai_dbi],
        errors="coerce"
    )

    data_dbi = data_dbi.dropna(subset=[kolom_k, kolom_nilai_dbi])

    if not data_dbi.empty:
        baris_terbaik = data_dbi.loc[data_dbi[kolom_nilai_dbi].idxmin()]
        k_optimal = int(baris_terbaik[kolom_k])
        dbi_terbaik = float(baris_terbaik[kolom_nilai_dbi])

# ============================================================
# 6. HEADER DENGAN GAMBAR RUMAH GADANG
# ============================================================

kolom_judul, kolom_gambar = st.columns([1.65, 1], gap="large")

with kolom_judul:
    st.markdown(
        """
        <div class="dashboard-header">
            <div class="eyebrow">SUMATERA BARAT • DATA ANALYTICS</div>
            <h1>Dashboard Klasterisasi Sosial Ekonomi</h1>
            <p>
                Eksplorasi karakteristik sosial ekonomi kabupaten/kota
                di Provinsi Sumatera Barat menggunakan algoritma
                K-Means dan evaluasi Davies-Bouldin Index (DBI).
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

with kolom_gambar:
    try:
        st.image(
            "https://commons.wikimedia.org/wiki/Special:FilePath/Rumah_Gadang_Suku_Minang.jpg",
            use_container_width=True
        )
        st.caption(
            "Rumah Gadang, rumah adat Minangkabau. "
            "Foto: Vina Iasha / Wikimedia Commons, CC BY-SA 4.0."
        )
    except Exception:
        st.info("Rumah Gadang — ikon budaya Sumatera Barat.")

st.markdown("")

# ============================================================
# 7. SIDEBAR FILTER
# ============================================================

with st.sidebar:
    st.markdown("## 🏠 Dashboard Sumbar")
    st.markdown("---")

    st.markdown("### Filter Data")

    semua_klaster = sorted(
        data_klaster["Klaster"].dropna().unique(),
        key=lambda x: int(x) if str(x).isdigit() else str(x)
    )

    klaster_dipilih = st.multiselect(
        "Pilih klaster",
        options=semua_klaster,
        default=semua_klaster,
        format_func=lambda x: f"Klaster {x}"
    )

    semua_wilayah = sorted(data_klaster["Kabupaten/Kota"].unique())

    wilayah_dipilih = st.multiselect(
        "Pilih kabupaten/kota",
        options=semua_wilayah,
        default=semua_wilayah
    )

    st.markdown("---")
    st.markdown("### Informasi Penelitian")
    st.caption("Metode: K-Means Clustering")
    st.caption("Evaluasi: Davies-Bouldin Index (DBI)")
    st.caption("Objek: Kabupaten/kota di Sumatera Barat")

    st.markdown("---")
    st.caption("Dashboard hasil penelitian")

data_tampil = data_klaster[
    data_klaster["Klaster"].isin(klaster_dipilih)
    & data_klaster["Kabupaten/Kota"].isin(wilayah_dipilih)
].copy()

# ============================================================
# 8. KARTU RINGKASAN
# ============================================================

jumlah_wilayah = data_tampil["Kabupaten/Kota"].nunique()
jumlah_klaster_tampil = data_tampil["Klaster"].nunique()

st.markdown(
    '<div class="section-heading">Ringkasan Hasil Klasterisasi</div>',
    unsafe_allow_html=True
)
st.markdown(
    '<div class="section-subtitle">Ikhtisar hasil pengelompokan wilayah berdasarkan indikator sosial ekonomi.</div>',
    unsafe_allow_html=True
)

kartu1, kartu2, kartu3, kartu4 = st.columns(4)

with kartu1:
    st.metric(
        label="Jumlah Wilayah",
        value=jumlah_wilayah,
        help="Jumlah kabupaten/kota sesuai filter yang dipilih."
    )

with kartu2:
    st.metric(
        label="Jumlah Indikator",
        value=9,
        help="Sembilan indikator sosial ekonomi yang digunakan dalam penelitian."
    )

with kartu3:
    st.metric(
        label="Jumlah Klaster",
        value=jumlah_klaster_tampil,
        help="Jumlah klaster yang tampil setelah filter diterapkan."
    )

with kartu4:
    st.metric(
        label="K Optimal (DBI)",
        value=k_optimal,
        delta=f"DBI = {format_angka(dbi_terbaik, 4)}" if dbi_terbaik is not None else None,
        help="Nilai K dengan DBI terkecil pada tabel evaluasi."
    )

st.markdown("")

# ============================================================
# 9. TAB DASHBOARD
# ============================================================

tab_hasil, tab_peta, tab_distribusi, tab_profil, tab_pca, tab_dbi = st.tabs(
    [
        "📋 Hasil Klaster",
        "🗺️ Peta Wilayah",
        "📊 Distribusi",
        "📈 Profil Indikator",
        "🔵 Visualisasi PCA",
        "🎯 Evaluasi DBI"
    ]
)

# ============================================================
# TAB 1: HASIL KLASTER
# ============================================================

with tab_hasil:
    st.markdown("### Hasil Pengelompokan Kabupaten/Kota")
    st.markdown(
        "Tabel berikut menampilkan kabupaten/kota beserta klaster "
        "hasil penerapan algoritma K-Means."
    )

    kolom_tabel = data_tampil.copy()

    def warna_baris(row):
        warna = WARNA_KLASTER.get(str(row.get("Klaster", "")), "#FFFFFF")
        return [
            "background-color: #FFFFFF; color: #243447;"
            if kolom != "Klaster"
            else f"background-color: {warna}; color: #172033; font-weight: bold;"
            for kolom in row.index
        ]

    st.dataframe(
        kolom_tabel.style.apply(warna_baris, axis=1),
        use_container_width=True,
        hide_index=True
    )

    csv_hasil = kolom_tabel.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="⬇️ Unduh Hasil Klaster (CSV)",
        data=csv_hasil,
        file_name="hasil_klaster_sumatera_barat.csv",
        mime="text/csv"
    )

    st.markdown(
        """
        <div class="info-box">
            <b>Catatan:</b> Warna pada kolom Klaster digunakan sebagai penanda
            visual. Warna tidak mengubah hasil perhitungan algoritma K-Means.
        </div>
        """,
        unsafe_allow_html=True
    )

# ============================================================
# TAB 2: PETA WILAYAH
# ============================================================

with tab_peta:
    st.markdown("### Peta Persebaran Klaster")
    st.markdown(
        "Peta memperlihatkan persebaran kabupaten/kota berdasarkan "
        "klaster yang dihasilkan dari analisis."
    )

    if not Path(FILE_ZIP_PETA).exists():
        st.warning(
            f"File peta tidak ditemukan: {FILE_ZIP_PETA}. "
            "Unggah ZIP shapefile ke repositori GitHub agar peta dapat ditampilkan."
        )
    elif data_tampil.empty:
        st.info("Tidak ada wilayah yang sesuai dengan filter saat ini.")
    else:
        try:
            with tempfile.TemporaryDirectory() as folder_sementara:
                with zipfile.ZipFile(FILE_ZIP_PETA, "r") as arsip:
                    arsip.extractall(folder_sementara)

                daftar_shp = list(Path(folder_sementara).rglob("*.shp"))

                if not daftar_shp:
                    st.warning(
                        "File .shp tidak ditemukan di dalam ZIP. "
                        "Pastikan ZIP berisi berkas shapefile yang lengkap."
                    )
                else:
                    path_shp = daftar_shp[0]
                    gdf = gpd.read_file(path_shp)

                    kolom_nama_peta = cari_kolom(
                        gdf,
                        ["kab_kota", "Kabupaten/Kota", "Kabupaten_Kota", "WADMKK", "NAMOBJ"]
                    )

                    if kolom_nama_peta is None:
                        st.warning(
                            "Kolom nama kabupaten/kota tidak ditemukan pada shapefile. "
                            "Periksa atribut nama wilayah di file peta."
                        )
                    else:
                        gdf = gdf.to_crs(epsg=4326)
                        gdf["_nama_wilayah"] = gdf[kolom_nama_peta].astype(str).str.strip()
                        gdf["_kunci"] = gdf["_nama_wilayah"].apply(normalisasi_nama)

                        data_peta = data_tampil.copy()
                        data_peta["_kunci"] = data_peta["Kabupaten/Kota"].apply(normalisasi_nama)

                        # Penyesuaian variasi penulisan nama wilayah
                        gdf["_kunci"] = gdf["_kunci"].replace({
                            "sawah lunto": "sawahlunto"
                        })
                        data_peta["_kunci"] = data_peta["_kunci"].replace({
                            "sawah lunto": "sawahlunto"
                        })

                        gdf_gabung = gdf.merge(
                            data_peta[["_kunci", "Kabupaten/Kota", "Klaster"]],
                            on="_kunci",
                            how="left"
                        )

                        gdf_gabung["Klaster"] = gdf_gabung["Klaster"].fillna("Tidak ada data")
                        gdf_gabung["_warna"] = gdf_gabung["Klaster"].map(WARNA_KLASTER)
                        gdf_gabung["_warna"] = gdf_gabung["_warna"].fillna("#D8DEE8")

                        geojson = json.loads(gdf_gabung.to_json())

                        fig_peta = go.Figure()

                        for nama_klaster in sorted(
                            gdf_gabung["Klaster"].dropna().unique(),
                            key=lambda x: int(x) if str(x).isdigit() else str(x)
                        ):
                            subset = gdf_gabung[gdf_gabung["Klaster"] == nama_klaster]

                            if subset.empty:
                                continue

                            geojson_subset = json.loads(subset.to_json())
                            daftar_lokasi = subset["_nama_wilayah"].tolist()

                            fig_peta.add_trace(
                                go.Choroplethmapbox(
                                    geojson=geojson_subset,
                                    locations=daftar_lokasi,
                                    z=[1] * len(daftar_lokasi),
                                    featureidkey="properties._nama_wilayah",
                                    colorscale=[
                                        [0, WARNA_KLASTER.get(str(nama_klaster), "#D8DEE8")],
                                        [1, WARNA_KLASTER.get(str(nama_klaster), "#D8DEE8")]
                                    ],
                                    showscale=False,
                                    marker_opacity=0.8,
                                    marker_line_width=1.2,
                                    marker_line_color="#FFFFFF",
                                    name=(
                                        f"Klaster {nama_klaster}"
                                        if str(nama_klaster).isdigit()
                                        else str(nama_klaster)
                                    ),
                                    hovertemplate=(
                                        "<b>%{location}</b><br>"
                                        f"Klaster: {nama_klaster}"
                                        "<extra></extra>"
                                    )
                                )
                            )

                        batas = gdf_gabung.total_bounds
                        tengah_lon = (batas[0] + batas[2]) / 2
                        tengah_lat = (batas[1] + batas[3]) / 2

                        fig_peta.update_layout(
                            mapbox={
                                "style": "open-street-map",
                                "center": {
                                    "lat": tengah_lat,
                                    "lon": tengah_lon
                                },
                                "zoom": 6.2
                            },
                            height=600,
                            margin={"r": 0, "t": 10, "l": 0, "b": 0},
                            legend={
                                "title": "Hasil Klaster",
                                "orientation": "h",
                                "yanchor": "bottom",
                                "y": 1.02,
                                "xanchor": "left",
                                "x": 0
                            },
                            paper_bgcolor="#FFFFFF",
                            plot_bgcolor="#FFFFFF"
                        )

                        st.plotly_chart(fig_peta, use_container_width=True)

                        st.markdown("#### Keterangan Warna")

                        k1, k2, k3 = st.columns(3)

                        with k1:
                            st.markdown("🟢 **Klaster 1**")

                        with k2:
                            st.markdown("🟡 **Klaster 2**")

                        with k3:
                            st.markdown("🔴 **Klaster 3**")

                        st.caption(
                            "Wilayah berwarna abu-abu menandakan data wilayah "
                            "belum cocok dengan nama pada shapefile atau tidak "
                            "termasuk dalam filter yang dipilih."
                        )

        except Exception as kesalahan:
            st.error(
                "Peta belum dapat ditampilkan. Periksa kelengkapan shapefile, "
                "nama kolom wilayah, serta dependensi aplikasi."
            )
            st.caption(f"Detail kesalahan: {kesalahan}")

# ============================================================
# TAB 3: DISTRIBUSI KLASTER
# ============================================================

with tab_distribusi:
    st.markdown("### Distribusi Kabupaten/Kota per Klaster")
    st.markdown(
        "Grafik berikut memperlihatkan jumlah wilayah yang termasuk "
        "ke dalam setiap klaster berdasarkan filter yang dipilih."
    )

    if data_tampil.empty:
        st.info("Tidak ada data untuk ditampilkan. Ubah filter pada sidebar.")
    else:
        distribusi = (
            data_tampil.groupby("Klaster")["Kabupaten/Kota"]
            .nunique()
            .reset_index(name="Jumlah Wilayah")
        )

        distribusi["Nama Klaster"] = distribusi["Klaster"].apply(
            lambda x: f"Klaster {x}"
        )

        distribusi["Warna"] = distribusi["Klaster"].map(WARNA_KLASTER)

        fig_distribusi = px.bar(
            distribusi,
            x="Nama Klaster",
            y="Jumlah Wilayah",
            text="Jumlah Wilayah",
            color="Klaster",
            color_discrete_map=WARNA_KLASTER,
            category_orders={
                "Klaster": sorted(distribusi["Klaster"].unique())
            },
            labels={
                "Nama Klaster": "Klaster",
                "Jumlah Wilayah": "Jumlah Kabupaten/Kota",
                "Klaster": "Klaster"
            }
        )

        fig_distribusi.update_traces(
            textposition="outside",
            cliponaxis=False
        )

        fig_distribusi.update_layout(
            height=420,
            showlegend=False,
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            margin={"l": 20, "r": 20, "t": 30, "b": 20},
            xaxis_title="",
            yaxis_title="Jumlah Kabupaten/Kota",
            font={"color": "#334155"}
        )

        st.plotly_chart(fig_distribusi, use_container_width=True)

        st.markdown("#### Rincian Wilayah")

        for klaster in sorted(
            data_tampil["Klaster"].unique(),
            key=lambda x: int(x) if str(x).isdigit() else str(x)
        ):
            daftar_wilayah = data_tampil.loc[
                data_tampil["Klaster"] == klaster,
                "Kabupaten/Kota"
            ].tolist()

            with st.expander(f"Klaster {klaster} — {len(daftar_wilayah)} wilayah"):
                st.write(", ".join(daftar_wilayah))

# ============================================================
# TAB 4: PROFIL INDIKATOR
# ============================================================

with tab_profil:
    st.markdown("### Profil Indikator Setiap Klaster")
    st.markdown(
        "Profil indikator dapat digunakan untuk membandingkan karakteristik "
        "masing-masing klaster berdasarkan nilai indikator yang tersedia "
        "pada sheet Profil_Klaster."
    )

    if data_profil.empty:
        st.warning(
            "Sheet Profil_Klaster belum tersedia atau kosong. "
            "Pastikan sheet tersebut terdapat dalam file Excel."
        )
    else:
        kolom_klaster_profil = cari_kolom(
            data_profil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        if kolom_klaster_profil is None:
            st.warning(
                "Kolom Klaster tidak ditemukan pada sheet Profil_Klaster."
            )
        else:
            data_profil_tampil = data_profil.copy()
            data_profil_tampil = data_profil_tampil.rename(
                columns={kolom_klaster_profil: "Klaster"}
            )
            data_profil_tampil["Klaster"] = (
                data_profil_tampil["Klaster"]
                .astype(str)
                .str.extract(r"(\d+)", expand=False)
            )

            data_profil_tampil = data_profil_tampil[
                data_profil_tampil["Klaster"].isin(klaster_dipilih)
            ]

            kolom_numerik = data_profil_tampil.select_dtypes(
                include="number"
            ).columns.tolist()

            # Kolom klaster tidak boleh dipakai sebagai indikator numerik
            kolom_numerik = [
                kolom for kolom in kolom_numerik
                if kolom != "Klaster"
            ]

            if data_profil_tampil.empty:
                st.info("Tidak ada profil klaster yang sesuai dengan filter.")
            elif not kolom_numerik:
                st.dataframe(
                    data_profil_tampil,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.dataframe(
                    data_profil_tampil,
                    use_container_width=True,
                    hide_index=True
                )

                indikator_dipilih = st.selectbox(
                    "Pilih indikator untuk dibandingkan",
                    options=kolom_numerik
                )

                fig_profil = px.bar(
                    data_profil_tampil,
                    x="Klaster",
                    y=indikator_dipilih,
                    color="Klaster",
                    color_discrete_map=WARNA_KLASTER,
                    text_auto=".3f",
                    labels={
                        "Klaster": "Klaster",
                        indikator_dipilih: indikator_dipilih
                    }
                )

                fig_profil.update_layout(
                    height=420,
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    showlegend=False,
                    xaxis_title="",
                    yaxis_title=indikator_dipilih,
                    font={"color": "#334155"}
                )

                st.plotly_chart(fig_profil, use_container_width=True)

                st.caption(
                    "Interpretasikan nilai indikator sesuai skala data pada "
                    "sheet Profil_Klaster. Nilai yang lebih tinggi tidak "
                    "selalu berarti kondisi yang lebih baik, karena arah "
                    "interpretasi bergantung pada indikator yang digunakan."
                )

# ============================================================
# TAB 5: VISUALISASI PCA
# ============================================================

with tab_pca:
    st.markdown("### Visualisasi Persebaran Wilayah dengan PCA")
    st.markdown(
        "PCA digunakan untuk membantu memvisualisasikan persebaran titik "
        "wilayah dalam dua dimensi. Penentuan klaster pada penelitian ini "
        "tetap dilakukan menggunakan algoritma K-Means."
    )

    if data_pca.empty:
        st.warning(
            "Sheet Koordinat_PCA belum tersedia atau kosong. "
            "Pastikan koordinat PCA sudah disimpan pada file Excel."
        )
    else:
        kolom_pc1 = cari_kolom(
            data_pca,
            ["PC1", "PCA1", "Principal Component 1", "Komponen 1"]
        )
        kolom_pc2 = cari_kolom(
            data_pca,
            ["PC2", "PCA2", "Principal Component 2", "Komponen 2"]
        )

        kolom_klaster_pca = cari_kolom(
            data_pca,
            ["Klaster", "Cluster", "Label Klaster"]
        )

        kolom_wilayah_pca = cari_kolom(
            data_pca,
            ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota", "Wilayah", "Daerah"]
        )

        if kolom_pc1 is None or kolom_pc2 is None:
            st.warning(
                "Kolom PC1 dan PC2 tidak ditemukan pada sheet Koordinat_PCA. "
                "Pastikan kolom koordinat PCA tersedia."
            )
        elif kolom_klaster_pca is None:
            st.warning(
                "Kolom klaster tidak ditemukan pada sheet Koordinat_PCA."
            )
        else:
            data_pca_tampil = data_pca.copy()

            data_pca_tampil[kolom_pc1] = pd.to_numeric(
                data_pca_tampil[kolom_pc1],
                errors="coerce"
            )
            data_pca_tampil[kolom_pc2] = pd.to_numeric(
                data_pca_tampil[kolom_pc2],
                errors="coerce"
            )

            data_pca_tampil[kolom_klaster_pca] = (
                data_pca_tampil[kolom_klaster_pca]
                .astype(str)
                .str.extract(r"(\d+)", expand=False)
            )

            data_pca_tampil = data_pca_tampil.dropna(
                subset=[kolom_pc1, kolom_pc2, kolom_klaster_pca]
            )

            data_pca_tampil = data_pca_tampil[
                data_pca_tampil[kolom_klaster_pca].isin(klaster_dipilih)
            ]

            if (
                kolom_wilayah_pca is not None
                and wilayah_dipilih
            ):
                data_pca_tampil = data_pca_tampil[
                    data_pca_tampil[kolom_wilayah_pca]
                    .astype(str)
                    .isin(wilayah_dipilih)
                ]

            if data_pca_tampil.empty:
                st.info("Tidak ada titik PCA yang sesuai dengan filter.")
            else:
                fig_pca = px.scatter(
                    data_pca_tampil,
                    x=kolom_pc1,
                    y=kolom_pc2,
                    color=kolom_klaster_pca,
                    color_discrete_map=WARNA_KLASTER,
                    hover_name=kolom_wilayah_pca
                    if kolom_wilayah_pca is not None
                    else None,
                    labels={
                        kolom_pc1: "Komponen Utama 1 (PC1)",
                        kolom_pc2: "Komponen Utama 2 (PC2)",
                        kolom_klaster_pca: "Klaster"
                    }
                )

                fig_pca.update_traces(
                    marker={"size": 13, "line": {"width": 1, "color": "#FFFFFF"}}
                )

                fig_pca.update_layout(
                    height=550,
                    paper_bgcolor="#FFFFFF",
                    plot_bgcolor="#FFFFFF",
                    font={"color": "#334155"},
                    legend_title_text="Klaster",
                    margin={"l": 20, "r": 20, "t": 30, "b": 20}
                )

                st.plotly_chart(fig_pca, use_container_width=True)

                st.markdown(
                    """
                    <div class="info-box">
                        <b>Catatan:</b> Visualisasi PCA menyajikan data dalam
                        dua dimensi untuk membantu melihat pola persebaran
                        wilayah. Jarak dan posisi titik pada grafik merupakan
                        representasi dua komponen utama, bukan dasar tunggal
                        untuk menentukan hasil klaster.
                    </div>
                    """,
                    unsafe_allow_html=True
                )

# ============================================================
# TAB 6: EVALUASI DBI
# ============================================================

with tab_dbi:
    st.markdown("### Evaluasi Jumlah Klaster Menggunakan DBI")
    st.markdown(
        "Davies-Bouldin Index (DBI) digunakan untuk mengevaluasi kualitas "
        "hasil klasterisasi. Secara umum, nilai DBI yang lebih kecil "
        "menunjukkan hasil pengelompokan yang lebih baik berdasarkan "
        "kekompakan dan pemisahan antarklaster."
    )

    if data_dbi.empty or kolom_k is None or kolom_nilai_dbi is None:
        st.warning(
            "Data evaluasi DBI tidak dapat ditampilkan. "
            "Periksa sheet Evaluasi_DBI dan pastikan terdapat kolom K dan DBI."
        )
    else:
        data_dbi_tampil = data_dbi.copy()

        fig_dbi = px.line(
            data_dbi_tampil,
            x=kolom_k,
            y=kolom_nilai_dbi,
            markers=True,
            labels={
                kolom_k: "Jumlah Klaster (K)",
                kolom_nilai_dbi: "Davies-Bouldin Index (DBI)"
            }
        )

        fig_dbi.update_traces(
            line={"color": BIRU_AKSI, "width": 3},
            marker={"size": 10, "color": BIRU_TUA}
        )

        if dbi_terbaik is not None:
            fig_dbi.add_trace(
                go.Scatter(
                    x=[k_optimal],
                    y=[dbi_terbaik],
                    mode="markers+text",
                    marker={
                        "size": 16,
                        "color": "#E45B5B",
                        "line": {"color": "#FFFFFF", "width": 2}
                    },
                    text=[f"K terbaik = {k_optimal}"],
                    textposition="top center",
                    name="DBI terendah",
                    hovertemplate=(
                        "K = %{x}<br>"
                        "DBI = %{y:.4f}"
                        "<extra></extra>"
                    )
                )
            )

        fig_dbi.update_layout(
            height=450,
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font={"color": "#334155"},
            margin={"l": 20, "r": 20, "t": 30, "b": 20},
            xaxis_title="Jumlah Klaster (K)",
            yaxis_title="Nilai DBI",
            legend_title_text=""
        )

        st.plotly_chart(fig_dbi, use_container_width=True)

        if dbi_terbaik is not None:
            st.markdown(
                f"""
                <div class="info-box">
                    <b>Hasil evaluasi:</b> berdasarkan tabel evaluasi DBI,
                    nilai DBI terendah diperoleh pada <b>K = {k_optimal}</b>
                    dengan nilai DBI sebesar <b>{format_angka(dbi_terbaik, 4)}</b>.
                    Nilai tersebut menjadi dasar pemilihan jumlah klaster
                    pada penelitian ini.
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown("#### Tabel Evaluasi DBI")

        st.dataframe(
            data_dbi_tampil,
            use_container_width=True,
            hide_index=True
        )

        csv_dbi = data_dbi_tampil.to_csv(index=False).encode("utf-8-sig")

        st.download_button(
            label="⬇️ Unduh Evaluasi DBI (CSV)",
            data=csv_dbi,
            file_name="evaluasi_dbi_sumatera_barat.csv",
            mime="text/csv"
        )

# ============================================================
# 10. FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer-box">
        <b>Dashboard Klasterisasi Daerah Sumatera Barat</b><br>
        Analisis indikator sosial ekonomi menggunakan K-Means
        dan evaluasi Davies-Bouldin Index (DBI).<br>
        Visualisasi ini disusun untuk mendukung penyajian hasil penelitian.
    </div>
    """,
    unsafe_allow_html=True
)
