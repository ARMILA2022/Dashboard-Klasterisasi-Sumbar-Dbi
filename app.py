import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import geopandas as gpd
import json
import os
import zipfile
import tempfile
import unicodedata
import re

# =========================================================
# 1. KONFIGURASI DASHBOARD
# =========================================================

st.set_page_config(
    page_title="Dashboard Sumbar",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

WARNA_KLASTER = {
    1: "#2E8B57",  # Hijau
    2: "#F2C94C",  # Kuning
    3: "#E74C3C"   # Merah
}

WARNA_BIRU = "#174A7E"
WARNA_BIRU_MUDA = "#EAF3FB"

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_SHAPEFILE_ZIP = "Provinsi Sumatera Barat-KAB_KOTA.zip"

SHEET_HASIL = "Hasil_Klaster"
SHEET_DBI = "Evaluasi_DBI"
SHEET_PROFIL = "Profil_Klaster"
SHEET_PCA = "Koordinat_PCA"

# =========================================================
# 2. TAMPILAN CSS
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
        min-height: 110px;
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

    .keterangan {
        color: #536579;
        font-size: 13px;
    }

    div[data-testid="stMetric"] {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #D9E5F2;
    }

    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# 3. FUNGSI PENDUKUNG
# =========================================================

def cari_kolom(data, daftar_nama):
    """Mencari nama kolom berdasarkan beberapa kemungkinan penulisan."""
    if data is None or data.empty:
        return None

    kolom = {str(c).strip().lower(): c for c in data.columns}

    for nama in daftar_nama:
        if nama.lower() in kolom:
            return kolom[nama.lower()]

    for nama in daftar_nama:
        for nama_asli in data.columns:
            if nama.lower() in str(nama_asli).strip().lower():
                return nama_asli

    return None


def normalisasi_nama(nama):
    """Menyamakan penulisan nama wilayah untuk pencocokan data dan peta."""
    if pd.isna(nama):
        return ""

    nama = str(nama).strip().lower()
    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter for karakter in nama
        if not unicodedata.combining(karakter)
    )

    nama = re.sub(r"\s+", " ", nama)
    nama = nama.replace("kabupaten ", "")
    nama = nama.replace("kab. ", "")
    nama = nama.replace("kab ", "")
    nama = nama.replace("kota sawah lunto", "kota sawahlunto")

    return nama.strip()


def bersihkan_klaster(data):
    """Menyamakan kolom hasil klaster menjadi Kabupaten/Kota dan Klaster."""
    data = data.copy()

    kolom_wilayah = cari_kolom(
        data,
        [
            "Kabupaten/Kota",
            "Kabupaten_Kota",
            "kab_kota",
            "nama",
            "wilayah",
            "daerah"
        ]
    )

    kolom_klaster = cari_kolom(
        data,
        ["Klaster", "Cluster", "Cluster_Label", "label_klaster"]
    )

    if kolom_wilayah is None:
        if data.index.name is not None:
            data = data.reset_index()
            kolom_wilayah = cari_kolom(
                data,
                ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota", "nama", "wilayah"]
            )

    if kolom_wilayah is None:
        raise ValueError(
            "Kolom nama Kabupaten/Kota tidak ditemukan pada sheet Hasil_Klaster."
        )

    if kolom_klaster is None:
        raise ValueError(
            "Kolom Klaster tidak ditemukan pada sheet Hasil_Klaster."
        )

    data = data.rename(
        columns={
            kolom_wilayah: "Kabupaten/Kota",
            kolom_klaster: "Klaster"
        }
    )

    data["Kabupaten/Kota"] = data["Kabupaten/Kota"].astype(str).str.strip()

    data["Klaster"] = (
        data["Klaster"]
        .astype(str)
        .str.extract(r"(\d+)", expand=False)
    )

    data["Klaster"] = pd.to_numeric(data["Klaster"], errors="coerce")
    data = data.dropna(subset=["Klaster"])
    data["Klaster"] = data["Klaster"].astype(int)

    data["_kunci"] = data["Kabupaten/Kota"].apply(normalisasi_nama)

    return data


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
    )
    st.session_state["pilihan_klaster"] = st.session_state.get(
        "semua_klaster", []
    )


def format_angka(nilai, desimal=3):
    try:
        return f"{float(nilai):,.{desimal}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "-"


# =========================================================
# 4. MEMBACA DATA EXCEL
# =========================================================

@st.cache_data
def muat_data_excel(lokasi_file):
    excel = pd.ExcelFile(lokasi_file)

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

    return hasil, evaluasi, profil, pca, excel.sheet_names


# =========================================================
# 5. MEMBACA SHAPEFILE
# =========================================================

@st.cache_data
def muat_shapefile(lokasi_zip):
    """Membaca shapefile kabupaten/kota dari berkas ZIP."""
    folder_sementara = tempfile.mkdtemp(prefix="sumbar_peta_")

    with zipfile.ZipFile(lokasi_zip, "r") as arsip:
        arsip.extractall(folder_sementara)

    berkas_shp = []

    for akar, _, daftar_file in os.walk(folder_sementara):
        for nama_file in daftar_file:
            if nama_file.lower().endswith(".shp"):
                berkas_shp.append(os.path.join(akar, nama_file))

    if not berkas_shp:
        raise FileNotFoundError(
            "Tidak ditemukan file .shp di dalam ZIP shapefile."
        )

    # Mengutamakan shapefile batas kabupaten/kota.
    pilihan_shp = [
        path for path in berkas_shp
        if "kab_kota" in os.path.basename(path).lower()
    ]

    lokasi_shp = pilihan_shp[0] if pilihan_shp else berkas_shp[0]

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
            "nama",
            "kabupaten"
        ]
    )

    if kolom_nama is None:
        kolom_teks = [
            kolom for kolom in peta.columns
            if kolom != peta.geometry.name
            and peta[kolom].dtype == "object"
        ]

        if kolom_teks:
            kolom_nama = kolom_teks[0]
        else:
            raise ValueError(
                "Kolom nama kabupaten/kota tidak ditemukan pada shapefile."
            )

    peta = peta.rename(columns={kolom_nama: "Nama_Peta"})
    peta["Nama_Peta"] = peta["Nama_Peta"].astype(str).str.strip()
    peta["_kunci"] = peta["Nama_Peta"].apply(normalisasi_nama)

    # Konversi ke GeoJSON untuk digunakan oleh Plotly.
    geojson = json.loads(peta.to_json())

    return peta, geojson


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
            Provinsi Sumatera Barat | Algoritma K-Means | Evaluasi Davies-Bouldin Index
        </p>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# =========================================================
# 7. PEMERIKSAAN FILE
# =========================================================

if not os.path.exists(FILE_EXCEL):
    st.error(
        f"File Excel '{FILE_EXCEL}' tidak ditemukan. "
        "Pastikan file Excel sudah diunggah ke repositori GitHub "
        "yang sama dengan app.py."
    )
    st.stop()

try:
    hasil_asli, evaluasi_dbi, profil, pca, daftar_sheet = muat_data_excel(
        FILE_EXCEL
    )
    data_hasil = bersihkan_klaster(hasil_asli)

except Exception as kesalahan:
    st.error(f"Data Excel gagal dibaca: {kesalahan}")
    st.info(
        "Periksa nama file Excel, nama sheet, serta kolom "
        "Kabupaten/Kota dan Klaster."
    )
    st.stop()

# =========================================================
# 8. MENYIAPKAN DATA EVALUASI DBI
# =========================================================

kolom_k = cari_kolom(evaluasi_dbi, ["K", "Jumlah Klaster", "Jumlah_Klaster"])
kolom_dbi = cari_kolom(evaluasi_dbi, ["DBI", "Davies-Bouldin Index", "DBI Score"])

if kolom_k is not None and kolom_dbi is not None:
    evaluasi_dbi = evaluasi_dbi.rename(
        columns={
            kolom_k: "K",
            kolom_dbi: "DBI"
        }
    )

    evaluasi_dbi["K"] = pd.to_numeric(
        evaluasi_dbi["K"], errors="coerce"
    )

    evaluasi_dbi["DBI"] = pd.to_numeric(
        evaluasi_dbi["DBI"], errors="coerce"
    )

    evaluasi_dbi = evaluasi_dbi.dropna(subset=["K", "DBI"])
    evaluasi_dbi["K"] = evaluasi_dbi["K"].astype(int)
    evaluasi_dbi = evaluasi_dbi.sort_values("K")

    if not evaluasi_dbi.empty:
        baris_terbaik = evaluasi_dbi.loc[evaluasi_dbi["DBI"].idxmin()]
        k_optimal = int(baris_terbaik["K"])
        dbi_terbaik = float(baris_terbaik["DBI"])
    else:
        k_optimal = 3
        dbi_terbaik = 0.6041
else:
    k_optimal = 3
    dbi_terbaik = 0.6041

# =========================================================
# 9. SIDEBAR: PENCARIAN DAN FILTER
# =========================================================

st.sidebar.markdown(
    """
    <h2 style="color:#174A7E;">🔎 Filter Dashboard</h2>
    """,
    unsafe_allow_html=True
)

semua_wilayah = sorted(data_hasil["Kabupaten/Kota"].unique().tolist())
semua_klaster = sorted(data_hasil["Klaster"].unique().tolist())

st.session_state["semua_wilayah"] = semua_wilayah
st.session_state["semua_klaster"] = semua_klaster

if "kata_cari" not in st.session_state:
    st.session_state["kata_cari"] = ""

if "pilihan_wilayah" not in st.session_state:
    st.session_state["pilihan_wilayah"] = semua_wilayah.copy()

if "pilihan_klaster" not in st.session_state:
    st.session_state["pilihan_klaster"] = semua_klaster.copy()

# Menghindari pilihan state yang tidak lagi tersedia.
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

for nomor, warna in WARNA_KLASTER.items():
    st.sidebar.markdown(
        f"""
        <div style="display:flex; align-items:center; margin-bottom:8px;">
            <div style="width:16px; height:16px; background:{warna};
                        border-radius:4px; margin-right:9px;"></div>
            <span>Klaster {nomor}</span>
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
        data_filter["Kabupaten/Kota"]
        .str.contains(kata_cari, case=False, na=False)
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
# 11. NAVIGASI HALAMAN
# =========================================================

halaman = st.sidebar.radio(
    "Navigasi",
    [
        "Beranda",
        "Hasil Klaster",
        "Peta Klaster",
        "Distribusi Klaster",
        "Profil Klaster",
        "Visualisasi PCA",
        "Evaluasi DBI"
    ]
)

# =========================================================
# 12. HALAMAN BERANDA
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
            "Jumlah Klaster",
            str(data_hasil["Klaster"].nunique())
        )

    with kolom4:
        tampilkan_kartu(
            "Nilai DBI Terbaik",
            format_angka(dbi_terbaik, 4)
        )

    st.markdown("### Tentang Dashboard")

    st.write(
        "Dashboard ini menyajikan hasil klasterisasi kabupaten/kota "
        "di Provinsi Sumatera Barat berdasarkan indikator sosial ekonomi "
        "menggunakan algoritma K-Means. Hasil klaster dapat digunakan "
        "untuk melihat kemiripan karakteristik antarwilayah."
    )

    st.write(
        "Penentuan jumlah klaster terbaik dilakukan dengan "
        "Davies-Bouldin Index (DBI). Nilai DBI yang lebih kecil "
        "menunjukkan hasil klaster yang lebih baik berdasarkan "
        "kekompakan klaster dan pemisahan antarklaster."
    )

    st.markdown("### Ringkasan Anggota Klaster")

    ringkasan = (
        data_hasil.groupby("Klaster")["Kabupaten/Kota"]
        .agg(["count", lambda x: ", ".join(sorted(x.astype(str)))])
        .reset_index()
    )

    ringkasan.columns = ["Klaster", "Jumlah Wilayah", "Anggota Wilayah"]

    st.dataframe(
        ringkasan,
        hide_index=True,
        width="stretch"
    )

    st.caption(
        "Gunakan menu Navigasi di sebelah kiri untuk melihat "
        "hasil, peta, profil indikator, PCA, dan evaluasi DBI."
    )

# =========================================================
# 13. HALAMAN HASIL KLASTER
# =========================================================

elif halaman == "Hasil Klaster":

    st.subheader("Hasil Klasterisasi Kabupaten/Kota")

    st.write(
        "Tabel berikut menampilkan hasil pengelompokan wilayah. "
        "Gunakan pencarian dan filter pada sidebar untuk mempersempit hasil."
    )

    st.metric("Wilayah yang Ditampilkan", len(data_filter))

    kolom_tampil = [
        kolom for kolom in data_filter.columns
        if kolom != "_kunci"
    ]

    st.dataframe(
        data_filter[kolom_tampil].sort_values("Kabupaten/Kota"),
        hide_index=True,
        width="stretch"
    )

    csv_hasil = data_filter[kolom_tampil].to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        "⬇️ Unduh Hasil Klaster (CSV)",
        data=csv_hasil,
        file_name="hasil_klaster_sumatera_barat.csv",
        mime="text/csv",
        use_container_width=True
    )

# =========================================================
# 14. HALAMAN PETA KLASTER
# =========================================================

elif halaman == "Peta Klaster":

    st.subheader("Peta Persebaran Klaster Kabupaten/Kota")

    st.write(
        "Warna pada peta menunjukkan klaster masing-masing wilayah: "
        "hijau untuk Klaster 1, kuning untuk Klaster 2, dan merah "
        "untuk Klaster 3."
    )

    if not os.path.exists(FILE_SHAPEFILE_ZIP):
        st.error(
            f"File '{FILE_SHAPEFILE_ZIP}' tidak ditemukan. "
            "Unggah ZIP shapefile kabupaten/kota Sumatera Barat "
            "ke repositori GitHub yang sama dengan app.py."
        )

    else:
        try:
            peta, geojson = muat_shapefile(FILE_SHAPEFILE_ZIP)

            gabungan = peta.merge(
                data_hasil[["_kunci", "Klaster", "Kabupaten/Kota"]],
                on="_kunci",
                how="left"
            )

            gabungan["Label_Klaster"] = gabungan["Klaster"].apply(
                lambda x: f"Klaster {int(x)}"
                if pd.notna(x) else "Tidak ada data"
            )

            # Gunakan kunci yang konsisten antara peta dan data.
            geojson_peta = json.loads(gabungan.to_json())

            # Pastikan fitur GeoJSON memakai kunci wilayah yang sama.
            for fitur in geojson_peta["features"]:
                props = fitur.get("properties", {})
                props["_kunci"] = normalisasi_nama(
                    props.get("Nama_Peta", "")
                )

            kunci_dipilih = set(data_filter["_kunci"].tolist())

            # Wilayah yang tidak dipilih tetap terlihat abu-abu pada peta.
            for fitur in geojson_peta["features"]:
                props = fitur.setdefault("properties", {})
                kunci = props.get("_kunci", "")

                baris_wilayah = data_hasil[
                    data_hasil["_kunci"] == kunci
                ]

                if kunci not in kunci_dipilih or baris_wilayah.empty:
                    props["Klaster_Tampilan"] = "Tidak dipilih"
                    props["Nama_Tampilan"] = props.get(
                        "Nama_Peta", kunci
                    )
                else:
                    nomor = int(baris_wilayah.iloc[0]["Klaster"])
                    props["Klaster_Tampilan"] = f"Klaster {nomor}"
                    props["Nama_Tampilan"] = baris_wilayah.iloc[0][
                        "Kabupaten/Kota"
                    ]

            warna_peta = {
                "Klaster 1": WARNA_KLASTER[1],
                "Klaster 2": WARNA_KLASTER[2],
                "Klaster 3": WARNA_KLASTER[3],
                "Tidak dipilih": "#D5DCE5"
            }

            fig_peta = px.choropleth_map(
                gabungan,
                geojson=geojson_peta,
                locations="_kunci",
                featureidkey="properties._kunci",
                color="Label_Klaster",
                hover_name="Nama_Peta",
                hover_data={
                    "Klaster": True,
                    "_kunci": False,
                    "Label_Klaster": False
                },
                color_discrete_map={
                    "Klaster 1": WARNA_KLASTER[1],
                    "Klaster 2": WARNA_KLASTER[2],
                    "Klaster 3": WARNA_KLASTER[3],
                    "Tidak ada data": "#D5DCE5"
                },
                map_style="open-street-map",
                center={"lat": -0.7399, "lon": 100.8000},
                zoom=6.2,
                opacity=0.65
            )

            # Samakan warna kategori dan tampilkan batas wilayah.
            fig_peta.update_traces(
                marker_line_color="white",
                marker_line_width=1
            )

            fig_peta.update_layout(
                height=650,
                margin={"r": 0, "t": 10, "l": 0, "b": 0},
                legend_title_text="Klaster",
                paper_bgcolor="white"
            )

            st.plotly_chart(
                fig_peta,
                width="stretch"
            )

            st.caption(
                "Peta menggunakan batas wilayah dari shapefile yang tersedia. "
                "Jika nama wilayah pada Excel berbeda dari shapefile, "
                "periksa kesesuaian nama wilayah tersebut."
            )

            daftar_cocok = set(peta["_kunci"].tolist())
            daftar_data = set(data_hasil["_kunci"].tolist())
            belum_cocok = sorted(daftar_data - daftar_cocok)

            if belum_cocok:
                st.warning(
                    "Nama wilayah berikut belum cocok dengan shapefile: "
                    + ", ".join(belum_cocok)
                )

        except Exception as kesalahan:
            st.error(f"Peta gagal ditampilkan: {kesalahan}")
            st.info(
                "Pastikan ZIP berisi shapefile lengkap (.shp, .shx, "
                ".dbf, dan .prj) serta memiliki kolom nama kabupaten/kota."
            )

# =========================================================
# 15. HALAMAN DISTRIBUSI KLASTER
# =========================================================

elif halaman == "Distribusi Klaster":

    st.subheader("Distribusi Jumlah Wilayah per Klaster")

    distribusi = (
        data_filter.groupby("Klaster")["Kabupaten/Kota"]
        .nunique()
        .reset_index(name="Jumlah Wilayah")
    )

    distribusi["Nama Klaster"] = distribusi["Klaster"].apply(
        lambda x: f"Klaster {x}"
    )

    distribusi["Warna"] = distribusi["Klaster"].map(WARNA_KLASTER)

    if distribusi.empty:
        st.warning("Tidak ada wilayah yang sesuai dengan filter.")
    else:
        fig_distribusi = px.bar(
            distribusi,
            x="Nama Klaster",
            y="Jumlah Wilayah",
            color="Nama Klaster",
            text="Jumlah Wilayah",
            color_discrete_map={
                f"Klaster {k}": v for k, v in WARNA_KLASTER.items()
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

            st.write(", ".join(anggota) if anggota else "-")

# =========================================================
# 16. HALAMAN PROFIL KLASTER
# =========================================================

elif halaman == "Profil Klaster":

    st.subheader("Profil Indikator Setiap Klaster")

    if profil.empty:
        st.warning(
            f"Sheet '{SHEET_PROFIL}' tidak tersedia atau tidak berisi data."
        )
    else:
        profil_tampil = profil.copy()

        kolom_klaster_profil = cari_kolom(
            profil_tampil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        if kolom_klaster_profil is None:
            # Jika label klaster tersimpan sebagai indeks.
            if (
                profil_tampil.columns.size > 0
                and str(profil_tampil.columns[0]).lower().startswith("unnamed")
            ):
                profil_tampil = profil_tampil.rename(
                    columns={profil_tampil.columns[0]: "Klaster"}
                )
            elif profil_tampil.index.name is not None:
                profil_tampil = profil_tampil.reset_index()
                kolom_klaster_profil = cari_kolom(
                    profil_tampil,
                    ["Klaster", "Cluster", "Nama Klaster"]
                )

        if kolom_klaster_profil is not None:
            profil_tampil = profil_tampil.rename(
                columns={kolom_klaster_profil: "Klaster"}
            )

        st.write(
            "Tabel dan grafik berikut menampilkan nilai indikator "
            "berdasarkan klaster sesuai data pada file Excel."
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
            c for c in kolom_numerik
            if c.lower() not in ["k", "jumlah wilayah", "count"]
        ]

        if kolom_numerik:
            indikator_pilihan = st.selectbox(
                "Pilih indikator untuk dibandingkan",
                options=kolom_numerik
            )

            kolom_label = cari_kolom(
                profil_tampil,
                ["Klaster", "Cluster", "Nama Klaster"]
            )

            if kolom_label is not None:
                fig_profil = px.bar(
                    profil_tampil,
                    x=kolom_label,
                    y=indikator_pilihan,
                    color=kolom_label,
                    title=f"Perbandingan {indikator_pilihan} antar Klaster"
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
            else:
                st.info(
                    "Kolom label klaster tidak ditemukan sehingga "
                    "grafik perbandingan tidak dapat dibuat."
                )
        else:
            st.info(
                "Tidak ditemukan kolom numerik yang dapat digunakan "
                "untuk membuat grafik profil."
            )

# =========================================================
# 17. HALAMAN VISUALISASI PCA
# =========================================================

elif halaman == "Visualisasi PCA":

    st.subheader("Visualisasi Persebaran Wilayah Menggunakan PCA")

    st.write(
        "PCA digunakan untuk memvisualisasikan persebaran wilayah "
        "dalam dua dimensi. Visualisasi ini membantu melihat posisi "
        "relatif wilayah berdasarkan hasil transformasi komponen utama, "
        "bukan sebagai dasar penentuan jumlah klaster."
    )

    if pca.empty:
        st.warning(
            f"Sheet '{SHEET_PCA}' tidak tersedia atau tidak berisi data."
        )
    else:
        pca_tampil = pca.copy()

        kolom_wilayah_pca = cari_kolom(
            pca_tampil,
            ["Kabupaten/Kota", "Kabupaten_Kota", "kab_kota", "nama", "wilayah"]
        )

        kolom_pc1 = cari_kolom(
            pca_tampil,
            ["PC1", "PCA1", "Komponen 1", "Principal Component 1"]
        )

        kolom_pc2 = cari_kolom(
            pca_tampil,
            ["PC2", "PCA2", "Komponen 2", "Principal Component 2"]
        )

        kolom_klaster_pca = cari_kolom(
            pca_tampil,
            ["Klaster", "Cluster", "Cluster_Label"]
        )

        if kolom_wilayah_pca is not None:
            pca_tampil = pca_tampil.rename(
                columns={kolom_wilayah_pca: "Kabupaten/Kota"}
            )

        if kolom_pc1 is not None:
            pca_tampil = pca_tampil.rename(columns={kolom_pc1: "PC1"})

        if kolom_pc2 is not None:
            pca_tampil = pca_tampil.rename(columns={kolom_pc2: "PC2"})

        if kolom_klaster_pca is not None:
            pca_tampil = pca_tampil.rename(
                columns={kolom_klaster_pca: "Klaster"}
            )

        if (
            "Kabupaten/Kota" in pca_tampil.columns
            and "PC1" in pca_tampil.columns
            and "PC2" in pca_tampil.columns
        ):
            pca_tampil["PC1"] = pd.to_numeric(
                pca_tampil["PC1"], errors="coerce"
            )

            pca_tampil["PC2"] = pd.to_numeric(
                pca_tampil["PC2"], errors="coerce"
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
                pca_tampil["Klaster"], errors="coerce"
            )

            pca_filter = pca_tampil[
                pca_tampil["_kunci"].isin(data_filter["_kunci"])
            ]

            if pca_filter.empty:
                st.warning(
                    "Tidak ada titik PCA yang sesuai dengan filter saat ini."
                )
            else:
                pca_filter = pca_filter.copy()
                pca_filter["Nama Klaster"] = pca_filter["Klaster"].apply(
                    lambda x: f"Klaster {int(x)}"
                    if pd.notna(x) else "Tidak diketahui"
                )

                warna_pca = {
                    f"Klaster {k}": v
                    for k, v in WARNA_KLASTER.items()
                }

                fig_pca = px.scatter(
                    pca_filter,
                    x="PC1",
                    y="PC2",
                    color="Nama Klaster",
                    color_discrete_map=warna_pca,
                    hover_name="Kabupaten/Kota",
                    title="Persebaran Kabupaten/Kota Berdasarkan Komponen Utama",
                    labels={
                        "PC1": "Komponen Utama 1 (PC1)",
                        "PC2": "Komponen Utama 2 (PC2)"
                    }
                )

                fig_pca.update_traces(
                    marker=dict(size=13, line=dict(width=1, color="white"))
                )

                fig_pca.update_layout(
                    height=580,
                    paper_bgcolor="white",
                    plot_bgcolor="white"
                )

                st.plotly_chart(
                    fig_pca,
                    width="stretch"
                )

                st.dataframe(
                    pca_filter.drop(columns=["_kunci", "Nama Klaster"]),
                    hide_index=True,
                    width="stretch"
                )

        else:
            st.error(
                "Kolom wilayah, PC1, atau PC2 tidak ditemukan pada sheet "
                "Koordinat_PCA. Periksa kembali nama kolom di file Excel."
            )

# =========================================================
# 18. HALAMAN EVALUASI DBI
# =========================================================

elif halaman == "Evaluasi DBI":

    st.subheader("Evaluasi Davies-Bouldin Index (DBI)")

    st.write(
        "Davies-Bouldin Index digunakan untuk mengevaluasi hasil "
        "klasterisasi. Nilai DBI yang lebih kecil menunjukkan "
        "hasil pengelompokan yang lebih baik berdasarkan kekompakan "
        "anggota dalam klaster dan pemisahan antarklaster."
    )

    kolom1, kolom2 = st.columns(2)

    with kolom1:
        st.metric(
            "Jumlah Klaster Terbaik",
            f"K = {k_optimal}"
        )

    with kolom2:
        st.metric(
            "Nilai DBI Terbaik",
            format_angka(dbi_terbaik, 4)
        )

    if evaluasi_dbi.empty:
        st.warning("Data evaluasi DBI tidak tersedia.")
    else:
        st.markdown("### Grafik Nilai DBI untuk Setiap K")

        fig_dbi = go.Figure()

        fig_dbi.add_trace(
            go.Scatter(
                x=evaluasi_dbi["K"],
                y=evaluasi_dbi["DBI"],
                mode="lines+markers",
                name="Nilai DBI",
                line=dict(color=WARNA_BIRU, width=3),
                marker=dict(size=9)
            )
        )

        fig_dbi.add_trace(
            go.Scatter(
                x=[k_optimal],
                y=[dbi_terbaik],
                mode="markers",
                name="DBI Terbaik",
                marker=dict(
                    color="#E74C3C",
                    size=15,
                    symbol="star"
                )
            )
        )

        fig_dbi.update_layout(
            xaxis_title="Jumlah Klaster (K)",
            yaxis_title="Davies-Bouldin Index",
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

        tabel_dbi["Status"] = tabel_dbi["K"].apply(
            lambda k: "K terbaik" if int(k) == k_optimal else ""
        )

        st.dataframe(
            tabel_dbi,
            hide_index=True,
            width="stretch"
        )

        csv_dbi = tabel_dbi.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ Unduh Evaluasi DBI (CSV)",
            data=csv_dbi,
            file_name="evaluasi_dbi_sumatera_barat.csv",
            mime="text/csv",
            use_container_width=True
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
        Provinsi Sumatera Barat | K-Means Clustering
    </div>
    """,
    unsafe_allow_html=True
)
