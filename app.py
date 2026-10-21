import os
import re
import json
import zipfile
import tempfile
import unicodedata

import numpy as np
import pandas as pd
import geopandas as gpd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# 1. KONFIGURASI DASHBOARD
# =========================================================

st.set_page_config(
    page_title="Dashboard Sumbar",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_SHAPEFILE = "Provinsi Sumatera Barat-KAB_KOTA.zip"

SHEET_HASIL = "Hasil_Klaster"
SHEET_DBI = "Evaluasi_DBI"
SHEET_PROFIL = "Profil_Klaster"
SHEET_PCA = "Koordinat_PCA"

WARNA_KLASTER = {
    "Klaster 1": "#2E8B57",
    "Klaster 2": "#F2C94C",
    "Klaster 3": "#E74C3C"
}

WARNA_BIRU = "#174A7E"


# =========================================================
# 2. TAMPILAN DASHBOARD
# =========================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #F5F8FC;
    }

    [data-testid="stSidebar"] {
        background-color: #EAF3FB;
    }

    h1, h2, h3 {
        color: #174A7E;
    }

    .kartu {
        background-color: white;
        padding: 18px;
        border-radius: 12px;
        border-left: 5px solid #174A7E;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        min-height: 105px;
    }

    .kartu-label {
        color: #536579;
        font-size: 14px;
        margin-bottom: 8px;
    }

    .kartu-nilai {
        color: #174A7E;
        font-size: 27px;
        font-weight: bold;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 3. FUNGSI PENDUKUNG
# =========================================================

def cari_kolom(data, daftar_nama):
    """Mencari kolom berdasarkan beberapa kemungkinan nama."""
    if data is None or data.empty:
        return None

    nama_kolom = {
        str(k).strip().lower(): k
        for k in data.columns
    }

    for nama in daftar_nama:
        if nama.strip().lower() in nama_kolom:
            return nama_kolom[nama.strip().lower()]

    for nama in daftar_nama:
        for kolom in data.columns:
            if nama.strip().lower() in str(kolom).strip().lower():
                return kolom

    return None


def normalisasi_nama(nama):
    """Menyamakan penulisan nama wilayah untuk pencocokan peta."""
    if pd.isna(nama):
        return ""

    nama = str(nama).strip().lower()

    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter
        for karakter in nama
        if not unicodedata.combining(karakter)
    )

    nama = re.sub(r"\s+", " ", nama)

    for awalan in [
        "kabupaten ",
        "kab. ",
        "kab "
    ]:
        if nama.startswith(awalan):
            nama = nama[len(awalan):]
            break

    nama = nama.replace(
        "kota sawah lunto",
        "kota sawahlunto"
    )

    nama = nama.replace(
        "sawah lunto",
        "sawahlunto"
    )

    return nama.strip()


def siapkan_hasil_klaster(data):
    """Merapikan kolom wilayah dan klaster."""
    data = data.copy()

    kolom_wilayah = cari_kolom(
        data,
        [
            "Kabupaten/Kota",
            "Kabupaten_Kota",
            "kab_kota",
            "Nama Wilayah",
            "Wilayah",
            "Daerah"
        ]
    )

    kolom_klaster = cari_kolom(
        data,
        [
            "Klaster",
            "Cluster",
            "Cluster_Label",
            "label_klaster"
        ]
    )

    if kolom_wilayah is None or kolom_klaster is None:
        raise ValueError(
            "Sheet Hasil_Klaster harus memiliki kolom nama "
            "Kabupaten/Kota dan Klaster."
        )

    data = data.rename(
        columns={
            kolom_wilayah: "Kabupaten/Kota",
            kolom_klaster: "Klaster"
        }
    )

    data["Kabupaten/Kota"] = (
        data["Kabupaten/Kota"].astype(str).str.strip()
    )

    data["Klaster"] = (
        data["Klaster"]
        .astype(str)
        .str.extract(r"(\d+)", expand=False)
    )

    data["Klaster"] = pd.to_numeric(
        data["Klaster"],
        errors="coerce"
    )

    data = data.dropna(
        subset=["Kabupaten/Kota", "Klaster"]
    )

    data["Klaster"] = data["Klaster"].astype(int)

    data["_kunci"] = data["Kabupaten/Kota"].apply(
        normalisasi_nama
    )

    data["Nama Klaster"] = data["Klaster"].apply(
        lambda x: f"Klaster {x}"
    )

    return data


def format_angka(nilai, desimal=4):
    try:
        hasil = f"{float(nilai):,.{desimal}f}"
        return hasil.replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "-"


def tampilkan_kartu(label, nilai):
    st.markdown(
        f"""
        <div class="kartu">
            <div class="kartu-label">{label}</div>
            <div class="kartu-nilai">{nilai}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def reset_filter():
    st.session_state["kata_cari"] = ""
    st.session_state["pilihan_wilayah"] = st.session_state.get(
        "semua_wilayah", []
    ).copy()
    st.session_state["pilihan_klaster"] = st.session_state.get(
        "semua_klaster", []
    ).copy()


# =========================================================
# 4. MEMBACA DATA EXCEL
# =========================================================

@st.cache_data
def muat_data_excel(lokasi_file):
    hasil = pd.read_excel(
        lokasi_file,
        sheet_name=SHEET_HASIL
    )

    evaluasi = pd.read_excel(
        lokasi_file,
        sheet_name=SHEET_DBI
    )

    try:
        profil = pd.read_excel(
            lokasi_file,
            sheet_name=SHEET_PROFIL
        )
    except Exception:
        profil = pd.DataFrame()

    try:
        pca = pd.read_excel(
            lokasi_file,
            sheet_name=SHEET_PCA
        )
    except Exception:
        pca = pd.DataFrame()

    return hasil, evaluasi, profil, pca


# =========================================================
# 5. MEMBACA SHAPEFILE
# =========================================================

@st.cache_data
def muat_peta(lokasi_zip):
    folder = tempfile.mkdtemp(prefix="peta_sumbar_")

    with zipfile.ZipFile(lokasi_zip, "r") as arsip:
        arsip.extractall(folder)

    daftar_shp = []

    for akar, _, berkas in os.walk(folder):
        for nama in berkas:
            if nama.lower().endswith(".shp"):
                daftar_shp.append(
                    os.path.join(akar, nama)
                )

    if not daftar_shp:
        raise FileNotFoundError(
            "Tidak ditemukan file SHP di dalam ZIP."
        )

    pilihan = [
        lokasi for lokasi in daftar_shp
        if "kab_kota" in os.path.basename(lokasi).lower()
    ]

    lokasi_shp = pilihan[0] if pilihan else daftar_shp[0]

    peta = gpd.read_file(lokasi_shp)

    if peta.crs is not None and peta.crs.to_epsg() != 4326:
        peta = peta.to_crs(epsg=4326)

    kolom_nama = cari_kolom(
        peta,
        [
            "kab_kota",
            "Kabupaten/Kota",
            "WADMKK",
            "NAMOBJ",
            "nama"
        ]
    )

    if kolom_nama is None:
        kolom_teks = [
            kolom for kolom in peta.columns
            if kolom != peta.geometry.name
            and peta[kolom].dtype == "object"
        ]

        if not kolom_teks:
            raise ValueError(
                "Kolom nama wilayah tidak ditemukan dalam SHP."
            )

        kolom_nama = kolom_teks[0]

    peta = peta.rename(
        columns={kolom_nama: "Nama_Peta"}
    )

    peta["Nama_Peta"] = (
        peta["Nama_Peta"].astype(str).str.strip()
    )

    peta["_kunci"] = peta["Nama_Peta"].apply(
        normalisasi_nama
    )

    return peta


# =========================================================
# 6. HEADER DASHBOARD
# =========================================================

kolom_gambar, kolom_judul = st.columns(
    [1, 5],
    vertical_alignment="center"
)

with kolom_gambar:
    st.image(
        "https://commons.wikimedia.org/wiki/Special:FilePath/Rumah_Gadang_Suku_Minang.jpg",
        width=125
    )

with kolom_judul:
    st.markdown(
        """
        <h1 style="color:#174A7E; margin-bottom:0;">
            🏠 Dashboard Sumbar
        </h1>

        <p style="font-size:19px; margin-top:5px; color:#536579;">
            Klasterisasi Indikator Sosial Ekonomi
        </p>

        <p style="font-size:13px; color:#65758B; margin-top:-8px;">
            Provinsi Sumatera Barat | K-Means Clustering
        </p>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")


# =========================================================
# 7. MEMUAT DATA
# =========================================================

if not os.path.exists(FILE_EXCEL):
    st.error(
        f"File '{FILE_EXCEL}' tidak ditemukan. "
        "Pastikan file Excel sudah diunggah ke repositori GitHub "
        "yang sama dengan app.py."
    )
    st.stop()

try:
    hasil_asli, evaluasi_dbi, profil, pca = muat_data_excel(
        FILE_EXCEL
    )

    data_hasil = siapkan_hasil_klaster(hasil_asli)

except Exception as error:
    st.error(f"Data gagal dibaca: {error}")
    st.info(
        "Periksa nama file Excel, nama sheet, dan nama kolom "
        "pada file Excel."
    )
    st.stop()


# =========================================================
# 8. MENYIAPKAN HASIL EVALUASI DBI
# =========================================================

kolom_k = cari_kolom(
    evaluasi_dbi,
    ["K", "Jumlah Klaster", "Jumlah_Klaster"]
)

kolom_dbi = cari_kolom(
    evaluasi_dbi,
    ["DBI", "Davies-Bouldin Index", "DBI Score"]
)

if kolom_k is None or kolom_dbi is None:
    st.error(
        "Kolom K dan DBI tidak ditemukan pada sheet Evaluasi_DBI."
    )
    st.stop()

evaluasi_dbi = evaluasi_dbi.rename(
    columns={
        kolom_k: "K",
        kolom_dbi: "DBI"
    }
)

evaluasi_dbi["K"] = pd.to_numeric(
    evaluasi_dbi["K"],
    errors="coerce"
)

evaluasi_dbi["DBI"] = pd.to_numeric(
    evaluasi_dbi["DBI"],
    errors="coerce"
)

evaluasi_dbi = evaluasi_dbi.dropna(
    subset=["K", "DBI"]
)

evaluasi_dbi["K"] = evaluasi_dbi["K"].astype(int)
evaluasi_dbi = evaluasi_dbi.sort_values("K")

if evaluasi_dbi.empty:
    st.error("Tidak ada data evaluasi DBI yang dapat ditampilkan.")
    st.stop()

baris_terbaik = evaluasi_dbi.loc[
    evaluasi_dbi["DBI"].idxmin()
]

K_OPTIMAL = int(baris_terbaik["K"])
DBI_TERBAIK = float(baris_terbaik["DBI"])


# =========================================================
# 9. SIDEBAR DAN FILTER
# =========================================================

st.sidebar.markdown(
    """
    <h2 style="color:#174A7E;">🔎 Filter Dashboard</h2>
    """,
    unsafe_allow_html=True
)

semua_wilayah = sorted(
    data_hasil["Kabupaten/Kota"].unique().tolist()
)

semua_klaster = sorted(
    data_hasil["Klaster"].unique().tolist()
)

st.session_state["semua_wilayah"] = semua_wilayah
st.session_state["semua_klaster"] = semua_klaster

if "kata_cari" not in st.session_state:
    st.session_state["kata_cari"] = ""

if "pilihan_wilayah" not in st.session_state:
    st.session_state["pilihan_wilayah"] = semua_wilayah.copy()

if "pilihan_klaster" not in st.session_state:
    st.session_state["pilihan_klaster"] = semua_klaster.copy()

st.session_state["pilihan_wilayah"] = [
    w for w in st.session_state["pilihan_wilayah"]
    if w in semua_wilayah
]

st.session_state["pilihan_klaster"] = [
    k for k in st.session_state["pilihan_klaster"]
    if k in semua_klaster
]

st.sidebar.text_input(
    "Cari kabupaten/kota",
    placeholder="Contoh: Agam atau Kota Padang",
    key="kata_cari"
)

st.sidebar.multiselect(
    "Pilih kabupaten/kota",
    options=semua_wilayah,
    key="pilihan_wilayah"
)

st.sidebar.multiselect(
    "Pilih klaster",
    options=semua_klaster,
    format_func=lambda k: f"Klaster {k}",
    key="pilihan_klaster"
)

st.sidebar.button(
    "↩️ Reset pencarian dan filter",
    on_click=reset_filter,
    use_container_width=True
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎨 Keterangan Klaster")

for nama, warna in WARNA_KLASTER.items():
    st.sidebar.markdown(
        f"""
        <div style="display:flex; align-items:center; margin-bottom:8px;">
            <div style="width:16px; height:16px;
                        background:{warna}; border-radius:4px;
                        margin-right:9px;"></div>
            <span>{nama}</span>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# 10. MENERAPKAN FILTER
# =========================================================

data_filter = data_hasil.copy()

kata_cari = st.session_state["kata_cari"].strip()

if kata_cari:
    data_filter = data_filter[
        data_filter["Kabupaten/Kota"].str.contains(
            kata_cari,
            case=False,
            na=False
        )
    ]

data_filter = data_filter[
    data_filter["Kabupaten/Kota"].isin(
        st.session_state["pilihan_wilayah"]
    )
]

data_filter = data_filter[
    data_filter["Klaster"].isin(
        st.session_state["pilihan_klaster"]
    )
]


# =========================================================
# 11. NAVIGASI MENU
# URUTAN SESUAI ALUR PENELITIAN
# =========================================================

urutan_halaman = [
    "Beranda",
    "Evaluasi DBI",
    "Hasil Klaster",
    "Distribusi Klaster",
    "Profil Klaster",
    "Visualisasi PCA",
    "Peta Klaster"
]

halaman = st.sidebar.radio(
    "Navigasi",
    urutan_halaman,
    index=0
)


# =========================================================
# 12. BERANDA
# =========================================================

if halaman == "Beranda":

    st.subheader("Ringkasan Dashboard")

    kolom1, kolom2, kolom3, kolom4 = st.columns(4)

    with kolom1:
        tampilkan_kartu(
            "Jumlah Wilayah",
            str(data_hasil["Kabupaten/Kota"].nunique())
        )

    with kolom2:
        tampilkan_kartu(
            "Jumlah Indikator",
            "9"
        )

    with kolom3:
        tampilkan_kartu(
            "Jumlah Klaster Terbaik",
            str(K_OPTIMAL)
        )

    with kolom4:
        tampilkan_kartu(
            "Nilai DBI Terbaik",
            format_angka(DBI_TERBAIK, 4)
        )

    st.markdown("### Tentang Dashboard")

    st.write(
        "Dashboard ini menyajikan hasil klasterisasi kabupaten/kota "
        "di Provinsi Sumatera Barat berdasarkan sembilan indikator "
        "sosial ekonomi menggunakan algoritma K-Means."
    )

    st.write(
        "Evaluasi Davies-Bouldin Index (DBI) digunakan untuk "
        "membandingkan hasil pengelompokan pada beberapa nilai K. "
        "Nilai DBI terendah dipilih sebagai jumlah klaster terbaik "
        "dalam penelitian ini."
    )

    st.info(
        f"Berdasarkan tabel evaluasi yang dimuat, jumlah klaster "
        f"terbaik adalah K = {K_OPTIMAL}, dengan nilai DBI "
        f"{format_angka(DBI_TERBAIK, 4)}."
    )

    st.markdown("### Ringkasan Anggota Klaster")

    ringkasan = (
        data_hasil.groupby("Nama Klaster")["Kabupaten/Kota"]
        .agg(
            Jumlah_Wilayah="nunique",
            Anggota=lambda x: ", ".join(sorted(x.astype(str)))
        )
        .reset_index()
        .rename(
            columns={
                "Jumlah_Wilayah": "Jumlah Wilayah",
                "Anggota": "Anggota Wilayah"
            }
        )
    )

    st.dataframe(
        ringkasan,
        hide_index=True,
        width="stretch"
    )

    st.caption(
        "Gunakan menu Navigasi di sebelah kiri untuk melihat "
        "evaluasi DBI, hasil klaster, profil, visualisasi PCA, dan peta."
    )


# =========================================================
# 13. EVALUASI DBI
# =========================================================

elif halaman == "Evaluasi DBI":

    st.subheader("Evaluasi Davies-Bouldin Index (DBI)")

    st.write(
        "Halaman ini menampilkan hasil evaluasi jumlah klaster "
        "K = 2 sampai K = 10, sesuai nilai yang tersedia pada "
        "tabel evaluasi. Jumlah klaster terbaik ditentukan dari "
        "nilai DBI terendah."
    )

    kolom1, kolom2 = st.columns(2)

    with kolom1:
        st.metric(
            "Jumlah Klaster Terbaik",
            f"K = {K_OPTIMAL}"
        )

    with kolom2:
        st.metric(
            "Nilai DBI Terendah",
            format_angka(DBI_TERBAIK, 4)
        )

    st.markdown("### Grafik Perbandingan Nilai DBI")

    fig_dbi = go.Figure()

    fig_dbi.add_trace(
        go.Scatter(
            x=evaluasi_dbi["K"],
            y=evaluasi_dbi["DBI"],
            mode="lines+markers",
            name="Nilai DBI",
            line=dict(
                color=WARNA_BIRU,
                width=3
            ),
            marker=dict(size=9)
        )
    )

    fig_dbi.add_trace(
        go.Scatter(
            x=[K_OPTIMAL],
            y=[DBI_TERBAIK],
            mode="markers",
            name="DBI Terendah",
            marker=dict(
                color="#E74C3C",
                size=15,
                symbol="star"
            )
        )
    )

    fig_dbi.update_layout(
        xaxis_title="Jumlah Klaster (K)",
        yaxis_title="Nilai Davies-Bouldin Index",
        xaxis=dict(
            tickmode="linear",
            dtick=1
        ),
        height=470,
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend_title_text="Keterangan"
    )

    st.plotly_chart(
        fig_dbi,
        width="stretch"
    )

    st.markdown("### Tabel Evaluasi DBI")

    tabel_dbi = evaluasi_dbi.copy()

    tabel_dbi["Keterangan"] = tabel_dbi["K"].apply(
        lambda nilai: "K terbaik"
        if int(nilai) == K_OPTIMAL
        else ""
    )

    st.dataframe(
        tabel_dbi,
        hide_index=True,
        width="stretch"
    )

    st.download_button(
        "⬇️ Unduh Evaluasi DBI (CSV)",
        data=tabel_dbi.to_csv(
            index=False
        ).encode("utf-8-sig"),
        file_name="evaluasi_dbi_sumatera_barat.csv",
        mime="text/csv",
        use_container_width=True
    )

    st.caption(
        "Semakin kecil nilai DBI, semakin baik kualitas klaster "
        "berdasarkan ukuran evaluasi ini."
    )


# =========================================================
# 14. HASIL KLASTER
# =========================================================

elif halaman == "Hasil Klaster":

    st.subheader("Hasil Klasterisasi Kabupaten/Kota")

    st.write(
        f"Halaman ini menampilkan hasil pengelompokan wilayah "
        f"berdasarkan K terbaik, yaitu K = {K_OPTIMAL}, "
        "sesuai hasil evaluasi DBI."
    )

    st.metric(
        "Jumlah Wilayah Sesuai Filter",
        len(data_filter)
    )

    kolom_tampil = [
        kolom for kolom in data_filter.columns
        if kolom not in ["_kunci", "Nama Klaster"]
    ]

    st.dataframe(
        data_filter[kolom_tampil].sort_values(
            "Kabupaten/Kota"
        ),
        hide_index=True,
        width="stretch"
    )

    st.download_button(
        "⬇️ Unduh Hasil Klaster (CSV)",
        data=data_filter[kolom_tampil].to_csv(
            index=False
        ).encode("utf-8-sig"),
        file_name="hasil_klaster_sumatera_barat.csv",
        mime="text/csv",
        use_container_width=True
    )


# =========================================================
# 15. DISTRIBUSI KLASTER
# =========================================================

elif halaman == "Distribusi Klaster":

    st.subheader("Distribusi Jumlah Wilayah per Klaster")

    distribusi = (
        data_filter.groupby("Nama Klaster")["Kabupaten/Kota"]
        .nunique()
        .reset_index(name="Jumlah Wilayah")
    )

    urutan_klaster = [
        f"Klaster {nomor}"
        for nomor in sorted(data_hasil["Klaster"].unique())
    ]

    distribusi["Nama Klaster"] = pd.Categorical(
        distribusi["Nama Klaster"],
        categories=urutan_klaster,
        ordered=True
    )

    distribusi = distribusi.sort_values("Nama Klaster")

    if distribusi.empty:
        st.warning(
            "Tidak ada wilayah yang sesuai dengan filter saat ini."
        )
    else:
        fig_distribusi = px.bar(
            distribusi,
            x="Nama Klaster",
            y="Jumlah Wilayah",
            color="Nama Klaster",
            text="Jumlah Wilayah",
            color_discrete_map=WARNA_KLASTER,
            category_orders={
                "Nama Klaster": urutan_klaster
            },
            labels={
                "Nama Klaster": "Klaster",
                "Jumlah Wilayah": "Jumlah Kabupaten/Kota"
            }
        )

        fig_distribusi.update_traces(
            textposition="outside"
        )

        fig_distribusi.update_layout(
            showlegend=False,
            height=450,
            paper_bgcolor="white",
            plot_bgcolor="white"
        )

        st.plotly_chart(
            fig_distribusi,
            width="stretch"
        )

    st.markdown("### Anggota Setiap Klaster")

    for nomor in sorted(data_filter["Klaster"].unique()):
        anggota = sorted(
            data_filter.loc[
                data_filter["Klaster"] == nomor,
                "Kabupaten/Kota"
            ].tolist()
        )

        st.markdown(
            f"**Klaster {nomor} — {len(anggota)} wilayah**"
        )

        st.write(
            ", ".join(anggota) if anggota else "-"
        )


# =========================================================
# 16. PROFIL KLASTER
# =========================================================

elif halaman == "Profil Klaster":

    st.subheader("Profil Indikator Setiap Klaster")

    if profil.empty:
        st.warning(
            f"Sheet '{SHEET_PROFIL}' tidak ditemukan "
            "atau tidak berisi data."
        )

    else:
        profil_tampil = profil.copy()

        kolom_label = cari_kolom(
            profil_tampil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        if kolom_label is None:
            kolom_indeks = cari_kolom(
                profil_tampil,
                ["Unnamed: 0", "index"]
            )

            if kolom_indeks is not None:
                profil_tampil = profil_tampil.rename(
                    columns={kolom_indeks: "Klaster"}
                )
            else:
                st.warning(
                    "Kolom label klaster tidak ditemukan. "
                    "Pastikan sheet Profil_Klaster memiliki "
                    "kolom Klaster."
                )

        kolom_label = cari_kolom(
            profil_tampil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        if kolom_label is not None:
            profil_tampil = profil_tampil.rename(
                columns={kolom_label: "Klaster"}
            )

        st.write(
            "Tabel ini menampilkan profil indikator setiap klaster "
            "sesuai data yang tersimpan pada file Excel."
        )

        st.dataframe(
            profil_tampil,
            hide_index=True,
            width="stretch"
        )

        kolom_numerik = profil_tampil.select_dtypes(
            include=np.number
        ).columns.tolist()

        kolom_numerik = [
            kolom for kolom in kolom_numerik
            if kolom.lower() not in [
                "k",
                "jumlah wilayah",
                "count"
            ]
        ]

        kolom_label = cari_kolom(
            profil_tampil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        if kolom_numerik and kolom_label is not None:
            indikator = st.selectbox(
                "Pilih indikator untuk dibandingkan",
                kolom_numerik
            )

            data_grafik = profil_tampil.copy()

            data_grafik[indikator] = pd.to_numeric(
                data_grafik[indikator],
                errors="coerce"
            )

            fig_profil = px.bar(
                data_grafik,
                x=kolom_label,
                y=indikator,
                color=kolom_label,
                title=f"Perbandingan {indikator} antar Klaster",
                color_discrete_map=WARNA_KLASTER
            )

            fig_profil.update_layout(
                height=430,
                paper_bgcolor="white",
                plot_bgcolor="white",
                showlegend=False
            )

            st.plotly_chart(
                fig_profil,
                width="stretch"
            )

        elif kolom_numerik:
            st.info(
                "Grafik profil belum dapat dibuat karena kolom "
                "label klaster tidak ditemukan."
            )


# =========================================================
# 17. VISUALISASI PCA
# =========================================================

elif halaman == "Visualisasi PCA":

    st.subheader("Visualisasi Persebaran Wilayah Menggunakan PCA")

    st.write(
        "Principal Component Analysis (PCA) digunakan untuk "
        "memvisualisasikan persebaran titik wilayah dalam dua dimensi. "
        "Pada penelitian ini, PCA berfungsi sebagai visualisasi, "
        "bukan sebagai dasar penentuan jumlah klaster."
    )

    if pca.empty:
        st.warning(
            f"Sheet '{SHEET_PCA}' tidak ditemukan "
            "atau tidak berisi data."
        )

    else:
        pca_tampil = pca.copy()

        kolom_wilayah = cari_kolom(
            pca_tampil,
            [
                "Kabupaten/Kota",
                "Kabupaten_Kota",
                "kab_kota",
                "Wilayah",
                "Nama"
            ]
        )

        kolom_pc1 = cari_kolom(
            pca_tampil,
            ["PC1", "PCA1", "Komponen 1", "Principal Component 1"]
        )

        kolom_pc2 = cari_kolom(
            pca_tampil,
            ["PC2", "PCA2", "Komponen 2", "Principal Component 2"]
        )

        kolom_klaster = cari_kolom(
            pca_tampil,
            ["Klaster", "Cluster", "Cluster_Label"]
        )

        if (
            kolom_wilayah is not None
            and kolom_pc1 is not None
            and kolom_pc2 is not None
        ):
            pca_tampil = pca_tampil.rename(
                columns={
                    kolom_wilayah: "Kabupaten/Kota",
                    kolom_pc1: "PC1",
                    kolom_pc2: "PC2"
                }
            )

            if kolom_klaster is not None:
                pca_tampil = pca_tampil.rename(
                    columns={kolom_klaster: "Klaster"}
                )

            pca_tampil["PC1"] = pd.to_numeric(
                pca_tampil["PC1"],
                errors="coerce"
            )

            pca_tampil["PC2"] = pd.to_numeric(
                pca_tampil["PC2"],
                errors="coerce"
            )

            pca_tampil = pca_tampil.dropna(
                subset=["PC1", "PC2"]
            )

            pca_tampil["_kunci"] = pca_tampil[
                "Kabupaten/Kota"
            ].apply(normalisasi_nama)

            if "Klaster" not in pca_tampil.columns:
                pca_tampil = pca_tampil.merge(
                    data_hasil[["_kunci", "Klaster"]],
                    on="_kunci",
                    how="left"
                )

            pca_tampil["Klaster"] = pd.to_numeric(
                pca_tampil["Klaster"],
                errors="coerce"
            )

            pca_tampil["Nama Klaster"] = pca_tampil[
                "Klaster"
            ].apply(
                lambda x: f"Klaster {int(x)}"
                if pd.notna(x)
                else "Tidak diketahui"
            )

            pca_filter = pca_tampil[
                pca_tampil["_kunci"].isin(
                    data_filter["_kunci"]
                )
            ].copy()

            if pca_filter.empty:
                st.warning(
                    "Tidak ada titik PCA yang sesuai dengan filter."
                )

            else:
                fig_pca = px.scatter(
                    pca_filter,
                    x="PC1",
                    y="PC2",
                    color="Nama Klaster",
                    color_discrete_map={
                        **WARNA_KLASTER,
                        "Tidak diketahui": "#999999"
                    },
                    hover_name="Kabupaten/Kota",
                    title="Persebaran Kabupaten/Kota Berdasarkan PCA",
                    labels={
                        "PC1": "Komponen Utama 1 (PC1)",
                        "PC2": "Komponen Utama 2 (PC2)"
                    }
                )

                fig_pca.update_traces(
                    marker=dict(
                        size=13,
                        line=dict(
                            width=1,
                            color="white"
                        )
                    )
                )

                fig_pca.update_layout(
                    height=570,
                    paper_bgcolor="white",
                    plot_bgcolor="white"
                )

                st.plotly_chart(
                    fig_pca,
                    width="stretch"
                )

                kolom_tampil = [
                    kolom for kolom in pca_filter.columns
                    if kolom not in ["_kunci", "Nama Klaster"]
                ]

                st.dataframe(
                    pca_filter[kolom_tampil],
                    hide_index=True,
                    width="stretch"
                )

        else:
            st.error(
                "Kolom wilayah, PC1, dan PC2 tidak ditemukan "
                "pada sheet Koordinat_PCA. Periksa nama kolom "
                "pada file Excel."
            )


# =========================================================
# 18. PETA KLASTER
# =========================================================

elif halaman == "Peta Klaster":

    st.subheader("Peta Persebaran Klaster Kabupaten/Kota")

    st.write(
        "Peta ini menampilkan persebaran geografis hasil klasterisasi "
        "kabupaten/kota di Provinsi Sumatera Barat."
    )

    if not os.path.exists(FILE_SHAPEFILE):

        st.error(
            f"File '{FILE_SHAPEFILE}' tidak ditemukan. "
            "Unggah ZIP shapefile ke repositori GitHub yang sama "
            "dengan app.py."
        )

    else:
        try:
            peta = muat_peta(FILE_SHAPEFILE)

            data_peta = peta.merge(
                data_hasil[
                    ["_kunci", "Kabupaten/Kota", "Klaster"]
                ],
                on="_kunci",
                how="left"
            )

            # Tandai wilayah yang lolos filter.
            kunci_dipilih = set(
                data_filter["_kunci"].tolist()
            )

            data_peta["Nama Klaster Peta"] = data_peta.apply(
                lambda baris: (
                    f"Klaster {int(baris['Klaster'])}"
                    if (
                        pd.notna(baris["Klaster"])
                        and baris["_kunci"] in kunci_dipilih
                    )
                    else "Tidak dipilih"
                ),
                axis=1
            )

            data_peta["Nama Wilayah Peta"] = data_peta.apply(
                lambda baris: (
                    baris["Kabupaten/Kota"]
                    if pd.notna(baris["Kabupaten/Kota"])
                    else baris["Nama_Peta"]
                ),
                axis=1
            )

            geojson_peta = json.loads(
                data_peta.to_json()
            )

            # Tambahkan kunci yang konsisten pada setiap fitur GeoJSON.
            for fitur in geojson_peta["features"]:
                props = fitur.setdefault("properties", {})
                props["_kunci"] = normalisasi_nama(
                    props.get("Nama_Peta", "")
                )

            fig_peta = px.choropleth_map(
                data_peta,
                geojson=geojson_peta,
                locations="_kunci",
                featureidkey="properties._kunci",
                color="Nama Klaster Peta",
                hover_name="Nama Wilayah Peta",
                hover_data={
                    "Klaster": True,
                    "_kunci": False,
                    "Nama Klaster Peta": True,
                    "Nama Wilayah Peta": False
                },
                color_discrete_map={
                    **WARNA_KLASTER,
                    "Tidak dipilih": "#D5DCE5"
                },
                category_orders={
                    "Nama Klaster Peta": [
                        "Klaster 1",
                        "Klaster 2",
                        "Klaster 3",
                        "Tidak dipilih"
                    ]
                },
                map_style="open-street-map",
                center={
                    "lat": -0.7399,
                    "lon": 100.8000
                },
                zoom=6.2,
                opacity=0.68
            )

            fig_peta.update_traces(
                marker_line_color="white",
                marker_line_width=1
            )

            fig_peta.update_layout(
                height=650,
                margin={
                    "r": 0,
                    "t": 10,
                    "l": 0,
                    "b": 0
                },
                legend_title_text="Keterangan Klaster",
                paper_bgcolor="white"
            )

            st.plotly_chart(
                fig_peta,
                width="stretch"
            )

            st.caption(
                "Wilayah yang tidak sesuai dengan filter ditampilkan "
                "dengan warna abu-abu. Jika ada nama wilayah yang tidak "
                "cocok, periksa penulisan nama di Excel dan shapefile."
            )

            wilayah_peta = set(peta["_kunci"].tolist())
            wilayah_excel = set(data_hasil["_kunci"].tolist())

            belum_cocok = sorted(
                wilayah_excel - wilayah_peta
            )

            if belum_cocok:
                st.warning(
                    "Nama wilayah berikut belum cocok dengan shapefile: "
                    + ", ".join(belum_cocok)
                )

        except Exception as error:
            st.error(
                f"Peta gagal ditampilkan: {error}"
            )

            st.info(
                "Pastikan ZIP berisi file SHP, SHX, DBF, dan PRJ, "
                "serta memiliki kolom nama kabupaten/kota."
            )


# =========================================================
# 19. FOOTER
# =========================================================

st.markdown("---")

st.markdown(
    """
    <div style="text-align:center; color:#65758B; font-size:13px;">
        Dashboard Klasterisasi Indikator Sosial Ekonomi
        <br>
        Provinsi Sumatera Barat | Algoritma K-Means
    </div>
    """,
    unsafe_allow_html=True
)
