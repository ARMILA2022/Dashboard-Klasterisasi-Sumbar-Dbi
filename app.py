
import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
import os
import re
import json
import zipfile
import unicodedata
import geopandas as gpd

# ==================================================
# PENGATURAN DASHBOARD
# ==================================================
st.set_page_config(
    page_title="Dashboard Klasterisasi Sumatera Barat",
    page_icon="📊",
    layout="wide"
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_ZIP_PETA = "Provinsi Sumatera Barat-KAB_KOTA.zip"
FOLDER_PETA = "peta_sumbar_dashboard"

# Warna klaster dibuat konsisten pada seluruh grafik dan peta
WARNA_KLASTER = {
    "1": "#2E8B57",  # Hijau
    "2": "#F1C40F",  # Kuning
    "3": "#E74C3C"   # Merah
}

# ==================================================
# DESAIN DASHBOARD
# ==================================================
st.markdown("""
<style>
    .stApp {
        background-color: #F4F7F5;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3 {
        color: #174C35;
    }

    .dashboard-header {
        background: linear-gradient(120deg, #123D2B, #23734D);
        padding: 32px 34px;
        border-radius: 18px;
        margin-bottom: 24px;
        box-shadow: 0 6px 20px rgba(18, 61, 43, 0.13);
    }

    .dashboard-header h1 {
        color: #FFFFFF !important;
        font-size: 31px;
        font-weight: 750;
        line-height: 1.3;
        margin: 0 0 10px 0;
    }

    .dashboard-header p {
        color: #E6F4EA !important;
        font-size: 15px;
        line-height: 1.7;
        margin: 0;
    }

    .section-heading {
        color: #174C35;
        font-size: 23px;
        font-weight: 750;
        margin-top: 12px;
        margin-bottom: 6px;
    }

    [data-testid="stMetric"] {
        background: #FFFFFF;
        padding: 20px 22px;
        border-radius: 15px;
        border: 1px solid #E0EAE3;
        border-top: 4px solid #278052;
        box-shadow: 0 4px 14px rgba(20, 50, 35, 0.05);
    }

    [data-testid="stMetricLabel"] {
        color: #587064;
        font-weight: 600;
    }

    [data-testid="stMetricValue"] {
        color: #174C35;
        font-weight: 750;
    }

    [data-testid="stTabs"] button {
        font-weight: 600;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #E0EAE3;
        border-radius: 10px;
        overflow: hidden;
    }

    [data-testid="stSidebar"] {
        background-color: #EDF4EF;
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 9px;
        font-weight: 600;
    }

    hr {
        border-color: #DDE7E0;
    }

    .small-note {
        color: #64776C;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)


# ==================================================
# MEMUAT DATA
# ==================================================
if not Path(FILE_EXCEL).exists():
    st.error(
        f"File {FILE_EXCEL} tidak ditemukan. "
        "Pastikan file Excel berada di repository yang sama dengan app.py."
    )
    st.stop()


@st.cache_data
def muat_data():
    hasil = pd.read_excel(
        FILE_EXCEL, sheet_name="Hasil_Klaster"
    )
    evaluasi = pd.read_excel(
        FILE_EXCEL, sheet_name="Evaluasi_DBI"
    )
    profil = pd.read_excel(
        FILE_EXCEL, sheet_name="Profil_Klaster", index_col=0
    )
    data_pca = pd.read_excel(
        FILE_EXCEL, sheet_name="Koordinat_PCA"
    )

    return hasil, evaluasi, profil, data_pca


hasil, evaluasi, profil, data_pca = muat_data()

kolom_wilayah = "Kabupaten/Kota"
kolom_klaster = "Klaster"

if (
    kolom_wilayah not in hasil.columns
    or kolom_klaster not in hasil.columns
):
    st.error(
        "Kolom Kabupaten/Kota atau Klaster tidak ditemukan "
        "pada sheet Hasil_Klaster."
    )
    st.stop()

if evaluasi.empty or "DBI" not in evaluasi.columns or "K" not in evaluasi.columns:
    st.error("Data evaluasi DBI tidak memiliki kolom K dan DBI yang diperlukan.")
    st.stop()

baris_terbaik = evaluasi.loc[evaluasi["DBI"].idxmin()]
k_optimal = int(baris_terbaik["K"])
dbi_terbaik = float(baris_terbaik["DBI"])

jumlah_indikator = len(profil.columns)
jumlah_wilayah = hasil[kolom_wilayah].nunique()
daftar_klaster = sorted(
    hasil[kolom_klaster].dropna().unique(),
    key=lambda x: str(x)
)


# ==================================================
# HEADER DASHBOARD
# ==================================================
st.markdown("""
<div class="dashboard-header">
    <h1>Dashboard Klasterisasi Sosial Ekonomi</h1>
    <p>
        Kabupaten/Kota di Provinsi Sumatera Barat<br>
        Pengelompokan wilayah menggunakan algoritma K-Means
        dengan penentuan jumlah klaster berdasarkan
        Davies-Bouldin Index (DBI).
    </p>
</div>
""", unsafe_allow_html=True)


# ==================================================
# KARTU RINGKASAN
# ==================================================
st.markdown(
    '<div class="section-heading">Ringkasan Hasil Analisis</div>',
    unsafe_allow_html=True
)

a, b, c, d = st.columns(4)

a.metric(
    label="Jumlah Wilayah",
    value=f"{jumlah_wilayah} Kabupaten/Kota"
)

b.metric(
    label="Indikator Sosial Ekonomi",
    value=f"{jumlah_indikator} Indikator"
)

c.metric(
    label="Jumlah Klaster Optimal",
    value=f"K = {k_optimal}"
)

d.metric(
    label="Nilai DBI Terbaik",
    value=f"{dbi_terbaik:.4f}"
)

st.markdown(
    '<p class="small-note">Periode data: 2021–2025 | '
    'Sumber: BPS Provinsi Sumatera Barat</p>',
    unsafe_allow_html=True
)

st.divider()


# ==================================================
# FUNGSI NORMALISASI NAMA WILAYAH UNTUK PETA
# ==================================================
def normalisasi_nama_wilayah(nama):
    nama = str(nama).strip().lower()
    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter for karakter in nama
        if not unicodedata.combining(karakter)
    )
    nama = re.sub(r"\s+", " ", nama)
    nama = nama.replace("sawah lunto", "sawahlunto")
    return nama.strip()


# ==================================================
# PETA KLASTER
# ==================================================
st.markdown(
    '<div class="section-heading">Peta Klaster Kabupaten/Kota</div>',
    unsafe_allow_html=True
)

st.write(
    "Peta berikut memperlihatkan persebaran klaster sosial ekonomi "
    "pada 19 kabupaten/kota di Provinsi Sumatera Barat."
)

try:
    if not Path(FILE_ZIP_PETA).exists():
        raise FileNotFoundError(
            f"File batas wilayah {FILE_ZIP_PETA} tidak ditemukan."
        )

    # Ambil nama wilayah dan klaster dari hasil pengolahan
    df_hasil_peta = hasil[
        [kolom_wilayah, kolom_klaster]
    ].copy()

    df_hasil_peta["nama_kunci"] = (
        df_hasil_peta[kolom_wilayah]
        .apply(normalisasi_nama_wilayah)
    )

    df_hasil_peta["Klaster_Peta"] = (
        df_hasil_peta[kolom_klaster].astype(str)
    )

    # Ekstrak shapefile
    os.makedirs(FOLDER_PETA, exist_ok=True)

    with zipfile.ZipFile(FILE_ZIP_PETA, "r") as arsip:
        arsip.extractall(FOLDER_PETA)

    file_shp_peta = None

    for akar, folder, daftar_file in os.walk(FOLDER_PETA):
        for nama_file in daftar_file:
            if nama_file.lower().endswith(".shp"):
                file_shp_peta = os.path.join(akar, nama_file)
                break
        if file_shp_peta:
            break

    if file_shp_peta is None:
        raise FileNotFoundError(
            "File shapefile (.shp) tidak ditemukan di dalam ZIP."
        )

    gdf_peta = gpd.read_file(file_shp_peta)

    if "kab_kota" not in gdf_peta.columns:
        raise ValueError(
            "Kolom kab_kota tidak ditemukan pada shapefile."
        )

    gdf_peta["nama_kunci"] = (
        gdf_peta["kab_kota"]
        .apply(normalisasi_nama_wilayah)
    )

    # Gabungkan data spasial dengan label klaster
    gdf_peta = gdf_peta.merge(
        df_hasil_peta[["nama_kunci", "Klaster_Peta"]],
        on="nama_kunci",
        how="left",
        validate="one_to_one"
    )

    wilayah_tidak_cocok = gdf_peta.loc[
        gdf_peta["Klaster_Peta"].isna(), "kab_kota"
    ].tolist()

    if wilayah_tidak_cocok:
        st.warning(
            "Beberapa nama wilayah belum cocok antara Excel dan shapefile: "
            + ", ".join(wilayah_tidak_cocok)
        )
    else:
        # Gunakan warna tetap untuk tiga klaster penelitian
        warna_peta = {
            "1": "#2E8B57",
            "2": "#F1C40F",
            "3": "#E74C3C"
        }

        # Tetap sediakan warna jika file berisi label klaster lain
        for label in gdf_peta["Klaster_Peta"].dropna().unique():
            if label not in warna_peta:
                warna_peta[label] = "#64748B"

        geojson_peta = json.loads(gdf_peta.to_json())

        fig_peta = px.choropleth(
            gdf_peta,
            geojson=geojson_peta,
            locations="nama_kunci",
            featureidkey="properties.nama_kunci",
            color="Klaster_Peta",
            color_discrete_map=warna_peta,
            category_orders={
                "Klaster_Peta": sorted(
                    gdf_peta["Klaster_Peta"].dropna().unique(),
                    key=lambda x: str(x)
                )
            },
            hover_name="kab_kota",
            hover_data={
                "nama_kunci": False,
                "Klaster_Peta": True
            },
            labels={"Klaster_Peta": "Klaster"}
        )

        fig_peta.update_traces(
            marker_line_color="#FFFFFF",
            marker_line_width=1.3
        )

        fig_peta.update_geos(
            fitbounds="locations",
            visible=False,
            bgcolor="rgba(0,0,0,0)"
        )

        fig_peta.update_layout(
            height=720,
            autosize=True,
            margin={"r": 10, "t": 35, "l": 10, "b": 10},
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend_title_text="Klaster",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="left",
                x=0,
                font=dict(size=13)
            ),
            font=dict(
                family="Arial",
                size=13,
                color="#263238"
            )
        )

        st.plotly_chart(
            fig_peta,
            use_container_width=True,
            config={
                "displayModeBar": True,
                "scrollZoom": True,
                "responsive": True
            }
        )

        st.caption(
            "Keterangan warna: hijau = Klaster 1, "
            "kuning = Klaster 2, merah = Klaster 3. "
            "Warna menunjukkan hasil pengelompokan K-Means."
        )

        with st.expander("Lihat daftar wilayah dan klasternya"):
            tabel_peta = (
                gdf_peta[["kab_kota", "Klaster_Peta"]]
                .rename(columns={
                    "kab_kota": "Kabupaten/Kota",
                    "Klaster_Peta": "Klaster"
                })
                .sort_values("Kabupaten/Kota")
            )
            st.dataframe(
                tabel_peta,
                use_container_width=True,
                hide_index=True
            )

except Exception as e:
    st.error(f"Peta belum dapat ditampilkan: {e}")


# ==================================================
# FILTER DAN BAGIAN ANALISIS
# ==================================================
st.divider()

st.markdown(
    '<div class="section-heading">Detail Hasil Analisis</div>',
    unsafe_allow_html=True
)

pilihan_klaster = st.multiselect(
    "Pilih klaster yang ingin ditampilkan",
    options=daftar_klaster,
    default=daftar_klaster,
    help="Filter ini memengaruhi tabel hasil, distribusi wilayah, dan visualisasi PCA."
)

hasil_filter = hasil[
    hasil[kolom_klaster].isin(pilihan_klaster)
]

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Hasil Klaster",
    "Distribusi Wilayah",
    "Profil Indikator",
    "Visualisasi PCA",
    "Evaluasi DBI"
])


# ==================================================
# TAB 1: HASIL KLASTER
# ==================================================
with tab1:
    st.subheader("Hasil Pengelompokan Kabupaten/Kota")

    st.write(
        "Tabel berikut menunjukkan klaster setiap kabupaten/kota "
        "berdasarkan sembilan indikator sosial ekonomi yang digunakan."
    )

    st.dataframe(
        hasil_filter,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        label="Unduh Hasil Klaster (CSV)",
        data=hasil_filter.to_csv(index=False).encode("utf-8-sig"),
        file_name="hasil_klaster_sumbar.csv",
        mime="text/csv"
    )


# ==================================================
# TAB 2: DISTRIBUSI WILAYAH
# ==================================================
with tab2:
    st.subheader("Jumlah Wilayah pada Setiap Klaster")

    jumlah = (
        hasil_filter.groupby(kolom_klaster)[kolom_wilayah]
        .nunique()
        .reset_index(name="Jumlah Wilayah")
    )

    jumlah["Klaster_Grafik"] = jumlah[kolom_klaster].astype(str)

    warna_grafik = {
        str(k): v for k, v in WARNA_KLASTER.items()
    }

    fig = px.bar(
        jumlah,
        x="Klaster_Grafik",
        y="Jumlah Wilayah",
        color="Klaster_Grafik",
        color_discrete_map=warna_grafik,
        text="Jumlah Wilayah",
        title="Distribusi Anggota Klaster",
        labels={
            "Klaster_Grafik": "Klaster",
            "Jumlah Wilayah": "Jumlah Wilayah"
        }
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False
    )

    fig.update_layout(
        height=430,
        showlegend=False,
        margin=dict(l=20, r=20, t=65, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Anggota Setiap Klaster")

    for k in daftar_klaster:
        anggota = hasil_filter.loc[
            hasil_filter[kolom_klaster] == k,
            kolom_wilayah
        ].tolist()

        if anggota:
            st.markdown(f"**Klaster {k} — {len(anggota)} wilayah**")
            st.write(", ".join(anggota))


# ==================================================
# TAB 3: PROFIL INDIKATOR
# ==================================================
with tab3:
    st.subheader("Profil Sosial Ekonomi Setiap Klaster")

    st.write(
        "Profil klaster menyajikan nilai rata-rata indikator "
        "untuk membantu memahami karakteristik setiap kelompok wilayah."
    )

    st.dataframe(
        profil,
        use_container_width=True
    )

    indikator_pilihan = st.selectbox(
        "Pilih indikator yang ingin dibandingkan",
        options=list(profil.columns)
    )

    data_grafik = profil[indikator_pilihan].reset_index()
    data_grafik.columns = ["Klaster", "Rata-rata"]
    data_grafik["Klaster_Grafik"] = data_grafik["Klaster"].astype(str)

    fig = px.bar(
        data_grafik,
        x="Klaster_Grafik",
        y="Rata-rata",
        color="Klaster_Grafik",
        color_discrete_map=warna_grafik,
        text_auto=".3f",
        title=f"Rata-rata {indikator_pilihan} per Klaster",
        labels={
            "Klaster_Grafik": "Klaster",
            "Rata-rata": "Nilai rata-rata"
        }
    )

    fig.update_layout(
        height=430,
        showlegend=False,
        margin=dict(l=20, r=20, t=65, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )

    st.plotly_chart(fig, use_container_width=True)


# ==================================================
# TAB 4: VISUALISASI PCA
# ==================================================
with tab4:
    st.subheader("Persebaran Titik Daerah Berdasarkan PCA")

    pca_filter = data_pca[
        data_pca["Klaster"].isin(pilihan_klaster)
    ].copy()

    pca_filter["Klaster_Grafik"] = (
        pca_filter["Klaster"].astype(str)
    )

    fig = px.scatter(
        pca_filter,
        x="PC1",
        y="PC2",
        color="Klaster_Grafik",
        color_discrete_map=warna_grafik,
        hover_name="Kabupaten/Kota",
        title="Persebaran Kabupaten/Kota Sumatera Barat",
        labels={
            "PC1": "Komponen Utama 1",
            "PC2": "Komponen Utama 2",
            "Klaster_Grafik": "Klaster"
        }
    )

    fig.update_traces(
        marker=dict(size=12, line=dict(width=1, color="white"))
    )

    fig.update_layout(
        height=570,
        margin=dict(l=20, r=20, t=65, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "PCA digunakan untuk visualisasi persebaran titik daerah "
        "dalam dua dimensi, bukan sebagai dasar pembentukan klaster."
    )


# ==================================================
# TAB 5: EVALUASI DBI
# ==================================================
with tab5:
    st.subheader("Evaluasi Davies-Bouldin Index")

    st.write(
        "Nilai DBI digunakan untuk membandingkan hasil klasterisasi "
        "pada beberapa nilai K. Nilai DBI yang lebih rendah menunjukkan "
        "hasil klaster yang lebih baik berdasarkan indeks ini."
    )

    st.dataframe(
        evaluasi,
        use_container_width=True,
        hide_index=True
    )

    fig = px.line(
        evaluasi.sort_values("K"),
        x="K",
        y="DBI",
        markers=True,
        title="Evaluasi DBI untuk K = 2 sampai K = 10",
        labels={
            "K": "Jumlah Klaster (K)",
            "DBI": "Davies-Bouldin Index"
        }
    )

    fig.add_scatter(
        x=[k_optimal],
        y=[dbi_terbaik],
        mode="markers",
        marker=dict(
            size=15,
            symbol="star",
            color="#E74C3C",
            line=dict(width=1, color="white")
        ),
        name="K Optimal"
    )

    fig.update_layout(
        height=470,
        margin=dict(l=20, r=20, t=65, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="white"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.success(
        f"Berdasarkan nilai DBI terendah pada tabel evaluasi, "
        f"K optimal adalah {k_optimal} dengan DBI {dbi_terbaik:.4f}."
    )


# ==================================================
# FOOTER
# ==================================================
st.divider()

st.markdown("""
<div style="
    text-align:center;
    padding:18px 10px;
    color:#64776C;
    font-size:13px;
">
    <b>Dashboard Klasterisasi Sosial Ekonomi Sumatera Barat</b><br>
    Sumber data: BPS Provinsi Sumatera Barat<br>
    Pengolahan data menggunakan algoritma K-Means dan evaluasi DBI.
</div>
""", unsafe_allow_html=True)

