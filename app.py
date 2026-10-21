
import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(
    page_title="Klasterisasi Sosial Ekonomi Sumatera Barat",
    page_icon="📊",
    layout="wide"
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"

st.title("Dashboard Klasterisasi Sosial Ekonomi")
st.subheader("Kabupaten/Kota di Provinsi Sumatera Barat")
st.caption(
    "Algoritma K-Means dengan penentuan jumlah klaster "
    "menggunakan Davies-Bouldin Index (DBI)."
)

if not Path(FILE_EXCEL).exists():
    st.error("File hasil pengolahan Excel belum ditemukan.")
    st.stop()

@st.cache_data
def muat_data():
    hasil = pd.read_excel(FILE_EXCEL, sheet_name="Hasil_Klaster")
    evaluasi = pd.read_excel(FILE_EXCEL, sheet_name="Evaluasi_DBI")
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

baris_terbaik = evaluasi.loc[evaluasi["DBI"].idxmin()]
k_optimal = int(baris_terbaik["K"])
dbi_terbaik = float(baris_terbaik["DBI"])

jumlah_indikator = len(profil.columns)
jumlah_wilayah = hasil[kolom_wilayah].nunique()
daftar_klaster = sorted(hasil[kolom_klaster].unique())

# Ringkasan utama
a, b, c, d = st.columns(4)
a.metric("Jumlah Wilayah", jumlah_wilayah)
b.metric("Jumlah Indikator", jumlah_indikator)
c.metric("K Optimal", k_optimal)
d.metric("DBI Terbaik", f"{dbi_terbaik:.4f}")

st.divider()

# Filter klaster
pilihan_klaster = st.multiselect(
    "Pilih klaster yang ingin ditampilkan",
    daftar_klaster,
    default=daftar_klaster
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

with tab1:
    st.subheader("Hasil Pengelompokan Kabupaten/Kota")
    st.dataframe(hasil_filter, use_container_width=True)

    st.download_button(
        "Unduh hasil klaster CSV",
        hasil_filter.to_csv(index=False).encode("utf-8-sig"),
        file_name="hasil_klaster_sumbar.csv",
        mime="text/csv"
    )

with tab2:
    st.subheader("Jumlah Wilayah pada Setiap Klaster")

    jumlah = (
        hasil.groupby("Klaster")["Kabupaten/Kota"]
        .count()
        .reset_index(name="Jumlah Wilayah")
    )

    fig = px.bar(
        jumlah,
        x="Klaster",
        y="Jumlah Wilayah",
        color="Klaster",
        text="Jumlah Wilayah",
        title="Distribusi Anggota Klaster"
    )
    st.plotly_chart(fig, use_container_width=True)

    for k in daftar_klaster:
        anggota = hasil.loc[
            hasil["Klaster"] == k, "Kabupaten/Kota"
        ].tolist()
        st.write(f"**Klaster {k}:** {', '.join(anggota)}")

with tab3:
    st.subheader("Profil Sosial Ekonomi Setiap Klaster")
    st.dataframe(profil, use_container_width=True)

    indikator_pilihan = st.selectbox(
        "Pilih indikator",
        list(profil.columns)
    )

    data_grafik = profil[indikator_pilihan].reset_index()
    data_grafik.columns = ["Klaster", "Rata-rata"]

    fig = px.bar(
        data_grafik,
        x="Klaster",
        y="Rata-rata",
        color="Klaster",
        text_auto=".3f",
        title=f"Rata-rata {indikator_pilihan} per Klaster"
    )
    st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.subheader("Persebaran Titik Daerah Berdasarkan PCA")

    pca_filter = data_pca[
        data_pca["Klaster"].isin(pilihan_klaster)
    ]

    fig = px.scatter(
        pca_filter,
        x="PC1",
        y="PC2",
        color="Klaster",
        hover_name="Kabupaten/Kota",
        title="Persebaran Kabupaten/Kota Sumatera Barat",
        labels={
            "PC1": "Komponen Utama 1",
            "PC2": "Komponen Utama 2"
        }
    )

    st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "PCA digunakan untuk visualisasi persebaran titik daerah, "
        "bukan sebagai dasar pembentukan klaster."
    )

with tab5:
    st.subheader("Evaluasi Davies-Bouldin Index")

    st.dataframe(evaluasi, use_container_width=True)

    fig = px.line(
        evaluasi,
        x="K",
        y="DBI",
        markers=True,
        title="Evaluasi DBI untuk K=2 sampai K=10"
    )

    fig.add_scatter(
        x=[k_optimal],
        y=[dbi_terbaik],
        mode="markers",
        marker=dict(size=14, symbol="star"),
        name="K Optimal"
    )

    st.plotly_chart(fig, use_container_width=True)

st.divider()
st.caption(
    "Sumber data: BPS Provinsi Sumatera Barat, "
    "diolah menggunakan algoritma K-Means."
)




# ==================================================
# PETA KLASTER KABUPATEN/KOTA SUMATERA BARAT
# ==================================================
import os
import re
import json
import zipfile
import unicodedata
import geopandas as gpd
import pandas as pd
import plotly.express as px
import streamlit as st

st.divider()
st.header("Peta Klaster Kabupaten/Kota Sumatera Barat")
st.write(
    "Peta menampilkan hasil pengelompokan 19 kabupaten/kota "
    "berdasarkan indikator sosial ekonomi."
)

FILE_EXCEL_PETA = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_ZIP_PETA = "Provinsi Sumatera Barat-KAB_KOTA.zip"
FOLDER_PETA = "peta_sumbar_dashboard"

def normalisasi_nama_wilayah(nama):
    nama = str(nama).strip().lower()
    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter for karakter in nama
        if not unicodedata.combining(karakter)
    )
    nama = re.sub(r"\\s+", " ", nama)
    nama = nama.replace("sawah lunto", "sawahlunto")
    return nama.strip()


try:
    # Baca hasil klaster
    df_hasil_peta = pd.read_excel(
        FILE_EXCEL_PETA,
        sheet_name="Hasil_Klaster"
    )

    # Cari kolom nama wilayah dan label klaster
    kolom_wilayah = next(
        (c for c in ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota"]
         if c in df_hasil_peta.columns),
        None
    )
    kolom_klaster = next(
        (c for c in ["Klaster", "Cluster", "cluster"]
         if c in df_hasil_peta.columns),
        None
    )

    if kolom_wilayah is None or kolom_klaster is None:
        raise ValueError(
            "Kolom wilayah atau Klaster tidak ditemukan pada sheet "
            "Hasil_Klaster. Periksa nama kolom file Excel."
        )

    df_hasil_peta = df_hasil_peta[
        [kolom_wilayah, kolom_klaster]
    ].copy()

    df_hasil_peta["nama_kunci"] = (
        df_hasil_peta[kolom_wilayah]
        .apply(normalisasi_nama_wilayah)
    )
    df_hasil_peta["Klaster"] = (
        df_hasil_peta[kolom_klaster].astype(str)
    )

    # Ekstrak ZIP batas wilayah
    if not os.path.exists(FOLDER_PETA):
        os.makedirs(FOLDER_PETA, exist_ok=True)

    if not os.path.exists(FILE_ZIP_PETA):
        raise FileNotFoundError(
            "ZIP batas wilayah tidak ditemukan: " + FILE_ZIP_PETA
        )

    with zipfile.ZipFile(FILE_ZIP_PETA, "r") as arsip:
        arsip.extractall(FOLDER_PETA)

    # Cari file SHP di dalam folder hasil ekstraksi
    file_shp_peta = None
    for akar, folder, daftar in os.walk(FOLDER_PETA):
        for nama in daftar:
            if nama.lower().endswith(".shp"):
                file_shp_peta = os.path.join(akar, nama)
                break
        if file_shp_peta:
            break

    if file_shp_peta is None:
        raise FileNotFoundError(
            "File .shp tidak ditemukan dalam ZIP batas wilayah."
        )

    gdf_peta = gpd.read_file(file_shp_peta)

    if "kab_kota" not in gdf_peta.columns:
        raise ValueError(
            "Kolom kab_kota tidak ditemukan pada shapefile."
        )

    gdf_peta["nama_kunci"] = (
        gdf_peta["kab_kota"].apply(normalisasi_nama_wilayah)
    )

    # Gabungkan geometri dan hasil klaster
    gdf_peta = gdf_peta.merge(
        df_hasil_peta[["nama_kunci", "Klaster"]],
        on="nama_kunci",
        how="left",
        validate="one_to_one"
    )

    wilayah_tidak_cocok = gdf_peta.loc[
        gdf_peta["Klaster"].isna(), "kab_kota"
    ].tolist()

    if wilayah_tidak_cocok:
        st.error(
            "Ada wilayah yang belum cocok dengan hasil klaster: "
            + ", ".join(wilayah_tidak_cocok)
        )
    else:
        # Warna otomatis untuk seluruh klaster
        warna_tersedia = [
            "#2E8B57", "#F1C40F", "#E74C3C",
            "#3498DB", "#9B59B6", "#E67E22",
            "#1ABC9C", "#7F8C8D", "#34495E"
        ]

        daftar_klaster = sorted(
            gdf_peta["Klaster"].unique(),
            key=lambda x: int(float(x))
        )

        warna_klaster = {
            klaster: warna_tersedia[i % len(warna_tersedia)]
            for i, klaster in enumerate(daftar_klaster)
        }

        # Konversi GeoDataFrame ke GeoJSON untuk Plotly
        geojson_peta = json.loads(gdf_peta.to_json())

        fig_peta = px.choropleth(
            gdf_peta,
            geojson=geojson_peta,
            locations="nama_kunci",
            featureidkey="properties.nama_kunci",
            color="Klaster",
            color_discrete_map=warna_klaster,
            hover_name="kab_kota",
            hover_data={"nama_kunci": False, "Klaster": True},
            labels={"Klaster": "Klaster"}
        )

        fig_peta.update_geos(
            fitbounds="locations",
            visible=False
        )
        fig_peta.update_layout(
            height=650,
            margin={"r": 0, "t": 20, "l": 0, "b": 0},
            legend_title_text="Klaster"
        )

        st.plotly_chart(fig_peta, use_container_width=True)

        st.caption(
            "Warna menunjukkan label klaster hasil K-Means. "
            "Peta menggunakan batas administrasi kabupaten/kota "
            "Sumatera Barat."
        )

        st.subheader("Daftar Wilayah pada Peta")
        st.dataframe(
            gdf_peta[["kab_kota", "Klaster"]]
            .rename(columns={
                "kab_kota": "Kabupaten/Kota"
            })
            .sort_values("Kabupaten/Kota"),
            use_container_width=True,
            hide_index=True
        )

except Exception as e:
    st.error(f"Peta belum dapat ditampilkan: {e}")
