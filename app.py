import os
import re
import json
import zipfile
import tempfile
import unicodedata

import numpy as np
import pandas as pd
import geopandas as gpd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# =========================================================
# 1. KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="Dashboard Sumbar",
    page_icon="🏠",
    layout="wide"
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_ZIP = "Provinsi Sumatera Barat-KAB_KOTA.zip"

SHEET_HASIL = "Hasil_Klaster"
SHEET_DBI = "Evaluasi_DBI"
SHEET_PROFIL = "Profil_Klaster"
SHEET_PCA = "Koordinat_PCA"

WARNA = {
    1: "#2E8B57",
    2: "#F2C94C",
    3: "#E74C3C"
}

BIRU = "#174A7E"


# =========================================================
# 2. DESAIN
# =========================================================

st.markdown(
    """
    <style>
    .stApp { background-color: #F5F8FC; }
    [data-testid="stSidebar"] { background-color: #EAF3FB; }
    h1, h2, h3 { color: #174A7E; }

    .kartu {
        background: white;
        padding: 18px;
        border-radius: 12px;
        border-left: 5px solid #174A7E;
        box-shadow: 0 2px 8px rgba(0,0,0,.07);
        min-height: 100px;
    }
    .label-kartu { color: #536579; font-size: 14px; }
    .nilai-kartu {
        color: #174A7E;
        font-size: 27px;
        font-weight: bold;
        margin-top: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 3. FUNGSI BANTUAN
# =========================================================

def normalisasi_kolom(nama):
    """Menyamakan format nama kolom untuk pencarian."""
    teks = str(nama).strip().lower()
    teks = teks.replace("_", " ")
    teks = re.sub(r"[^a-z0-9 ]", " ", teks)
    return re.sub(r"\s+", " ", teks).strip()


def cari_kolom(data, pilihan):
    """Mencari nama kolom tanpa bergantung pada huruf besar/kecil."""
    if data is None or len(data.columns) == 0:
        return None

    kolom_asli = list(data.columns)
    kolom_normal = {
        normalisasi_kolom(k): k for k in kolom_asli
    }

    for nama in pilihan:
        kunci = normalisasi_kolom(nama)
        if kunci in kolom_normal:
            return kolom_normal[kunci]

    # Pencarian sebagian hanya jika pencarian persis gagal.
    for nama in pilihan:
        kunci = normalisasi_kolom(nama)
        if len(kunci) < 4:
            continue
        for kolom in kolom_asli:
            if kunci in normalisasi_kolom(kolom):
                return kolom

    return None


def normalisasi_wilayah(nama):
    if pd.isna(nama):
        return ""

    teks = str(nama).strip().lower()
    teks = unicodedata.normalize("NFKD", teks)
    teks = "".join(
        c for c in teks if not unicodedata.combining(c)
    )
    teks = re.sub(r"\s+", " ", teks)

    # Samakan penulisan Kota Sawah Lunto.
    teks = teks.replace("kota sawah lunto", "kota sawahlunto")
    teks = teks.replace("sawah lunto", "sawahlunto")

    # Hilangkan awalan kabupaten saja, bukan awalan kota.
    for awalan in ("kabupaten ", "kab. ", "kab "):
        if teks.startswith(awalan):
            teks = teks[len(awalan):]
            break

    return teks.strip()


def angka_numerik(kolom):
    """Mengubah kolom menjadi angka tanpa mengubah nilai yang valid."""
    return pd.to_numeric(kolom, errors="coerce")


def tampilkan_kartu(judul, nilai):
    st.markdown(
        f"""
        <div class="kartu">
            <div class="label-kartu">{judul}</div>
            <div class="nilai-kartu">{nilai}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def siapkan_hasil(data):
    data = data.copy()

    kolom_wilayah = cari_kolom(
        data,
        [
            "Kabupaten/Kota",
            "Kabupaten Kota",
            "kab_kota",
            "Nama Wilayah",
            "Wilayah",
            "Daerah"
        ]
    )

    kolom_klaster = cari_kolom(
        data,
        ["Klaster", "Cluster", "Cluster Label", "Label Klaster"]
    )

    if kolom_wilayah is None or kolom_klaster is None:
        raise ValueError(
            "Kolom Kabupaten/Kota atau Klaster tidak ditemukan "
            "pada sheet Hasil_Klaster. Kolom yang terbaca: "
            + ", ".join(map(str, data.columns))
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
        data["Klaster"].astype(str)
        .str.extract(r"(\d+)", expand=False)
    )

    data["Klaster"] = pd.to_numeric(
        data["Klaster"], errors="coerce"
    )

    data = data.dropna(
        subset=["Kabupaten/Kota", "Klaster"]
    ).copy()

    data["Klaster"] = data["Klaster"].astype(int)
    data["_kunci"] = data["Kabupaten/Kota"].apply(
        normalisasi_wilayah
    )

    data["Nama Klaster"] = data["Klaster"].apply(
        lambda x: f"Klaster {x}"
    )

    return data


# =========================================================
# 4. MEMBACA EXCEL
# =========================================================

@st.cache_data
def baca_excel(nama_file):
    hasil = pd.read_excel(
        nama_file,
        sheet_name=SHEET_HASIL
    )

    evaluasi = pd.read_excel(
        nama_file,
        sheet_name=SHEET_DBI
    )

    try:
        profil = pd.read_excel(
            nama_file,
            sheet_name=SHEET_PROFIL
        )
    except Exception:
        profil = pd.DataFrame()

    try:
        pca = pd.read_excel(
            nama_file,
            sheet_name=SHEET_PCA
        )
    except Exception:
        pca = pd.DataFrame()

    return hasil, evaluasi, profil, pca


# =========================================================
# 5. MENYIAPKAN PETA
# =========================================================

@st.cache_data
def baca_shapefile(nama_zip):
    folder = tempfile.mkdtemp(prefix="sumbar_")

    with zipfile.ZipFile(nama_zip, "r") as arsip:
        arsip.extractall(folder)

    daftar_shp = []

    for akar, _, berkas in os.walk(folder):
        for nama in berkas:
            if nama.lower().endswith(".shp"):
                daftar_shp.append(
                    os.path.join(akar, nama)
                )

    if not daftar_shp:
        raise ValueError(
            "File ZIP tidak berisi shapefile .shp."
        )

    pilihan = [
        f for f in daftar_shp
        if "kab_kota" in os.path.basename(f).lower()
    ]

    lokasi_shp = pilihan[0] if pilihan else daftar_shp[0]
    peta = gpd.read_file(lokasi_shp)

    if peta.crs is not None and peta.crs.to_epsg() != 4326:
        peta = peta.to_crs(epsg=4326)

    kolom_nama = cari_kolom(
        peta,
        ["kab_kota", "Kabupaten/Kota", "WADMKK", "NAMOBJ", "Nama"]
    )

    if kolom_nama is None:
        kolom_teks = [
            c for c in peta.columns
            if c != peta.geometry.name
            and peta[c].dtype == "object"
        ]

        if not kolom_teks:
            raise ValueError(
                "Kolom nama kabupaten/kota tidak ditemukan pada SHP."
            )

        kolom_nama = kolom_teks[0]

    peta = peta.rename(
        columns={kolom_nama: "Nama_Peta"}
    )

    peta["Nama_Peta"] = peta["Nama_Peta"].astype(str).str.strip()
    peta["_kunci"] = peta["Nama_Peta"].apply(
        normalisasi_wilayah
    )

    # Pastikan setiap fitur GeoJSON memiliki kunci pencocokan.
    geojson = json.loads(peta.to_json())

    for fitur in geojson["features"]:
        properti = fitur.setdefault("properties", {})
        properti["_kunci"] = normalisasi_wilayah(
            properti.get("Nama_Peta", "")
        )

    return peta, geojson


# =========================================================
# 6. HEADER
# =========================================================

kolom_logo, kolom_judul = st.columns(
    [1, 5],
    vertical_alignment="center"
)

with kolom_logo:
    st.image(
        "https://commons.wikimedia.org/wiki/Special:FilePath/Rumah_Gadang_Suku_Minang.jpg",
        width=125
    )

with kolom_judul:
    st.markdown(
        """
        <h1 style="margin-bottom:0;">🏠 Dashboard Sumbar</h1>
        <p style="font-size:19px; color:#536579; margin-top:5px;">
            Klasterisasi Indikator Sosial Ekonomi
        </p>
        <p style="font-size:13px; color:#65758B;">
            Provinsi Sumatera Barat | Algoritma K-Means
        </p>
        """,
        unsafe_allow_html=True
    )

st.divider()


# =========================================================
# 7. MEMUAT DATA UTAMA
# =========================================================

if not os.path.exists(FILE_EXCEL):
    st.error(
        f"File Excel '{FILE_EXCEL}' tidak ditemukan. "
        "Pastikan nama file di GitHub sama persis."
    )
    st.stop()

try:
    hasil_asli, evaluasi_dbi, profil, pca = baca_excel(FILE_EXCEL)
    hasil = siapkan_hasil(hasil_asli)

except Exception as error:
    st.error(f"Data Excel gagal dibaca: {error}")
    st.info(
        "Periksa nama file, nama sheet, serta nama kolom "
        "di file Excel."
    )
    st.stop()


# =========================================================
# 8. MEMBACA EVALUASI DBI DENGAN NAMA KOLOM YANG TEPAT
# =========================================================

kolom_k = cari_kolom(
    evaluasi_dbi,
    ["Jumlah Klaster", "Jumlah_Klaster", "K"]
)

kolom_dbi = cari_kolom(
    evaluasi_dbi,
    ["Davies-Bouldin Index", "DBI Score", "DBI"]
)

if kolom_k is None or kolom_dbi is None:
    st.error(
        "Kolom K atau DBI tidak ditemukan pada sheet Evaluasi_DBI. "
        "Kolom yang terbaca: "
        + ", ".join(map(str, evaluasi_dbi.columns))
    )
    st.stop()

evaluasi_dbi = evaluasi_dbi.rename(
    columns={
        kolom_k: "K",
        kolom_dbi: "DBI"
    }
)

evaluasi_dbi["K"] = angka_numerik(evaluasi_dbi["K"])
evaluasi_dbi["DBI"] = angka_numerik(evaluasi_dbi["DBI"])

evaluasi_dbi = evaluasi_dbi.dropna(
    subset=["K", "DBI"]
).copy()

evaluasi_dbi["K"] = evaluasi_dbi["K"].astype(int)
evaluasi_dbi = evaluasi_dbi.sort_values("K")

if evaluasi_dbi.empty:
    st.error("Tabel evaluasi DBI tidak memiliki nilai numerik.")
    st.stop()

baris_terbaik = evaluasi_dbi.loc[
    evaluasi_dbi["DBI"].idxmin()
]

K_OPTIMAL = int(baris_terbaik["K"])
DBI_TERBAIK = float(baris_terbaik["DBI"])


# =========================================================
# 9. SIDEBAR DAN FILTER
# =========================================================

st.sidebar.title("🔎 Filter Dashboard")

daftar_wilayah = sorted(
    hasil["Kabupaten/Kota"].unique().tolist()
)

daftar_klaster = sorted(
    hasil["Klaster"].unique().tolist()
)

st.session_state["semua_wilayah"] = daftar_wilayah
st.session_state["semua_klaster"] = daftar_klaster

if "kata_cari" not in st.session_state:
    st.session_state["kata_cari"] = ""

if "wilayah_pilihan" not in st.session_state:
    st.session_state["wilayah_pilihan"] = daftar_wilayah.copy()

if "klaster_pilihan" not in st.session_state:
    st.session_state["klaster_pilihan"] = daftar_klaster.copy()

st.session_state["wilayah_pilihan"] = [
    x for x in st.session_state["wilayah_pilihan"]
    if x in daftar_wilayah
]

st.session_state["klaster_pilihan"] = [
    x for x in st.session_state["klaster_pilihan"]
    if x in daftar_klaster
]

st.sidebar.text_input(
    "Cari kabupaten/kota",
    placeholder="Contoh: Agam",
    key="kata_cari"
)

st.sidebar.multiselect(
    "Pilih wilayah",
    options=daftar_wilayah,
    key="wilayah_pilihan"
)

st.sidebar.multiselect(
    "Pilih klaster",
    options=daftar_klaster,
    format_func=lambda x: f"Klaster {x}",
    key="klaster_pilihan"
)

def reset_filter():
    st.session_state["kata_cari"] = ""
    st.session_state["wilayah_pilihan"] = st.session_state[
        "semua_wilayah"
    ].copy()
    st.session_state["klaster_pilihan"] = st.session_state[
        "semua_klaster"
    ].copy()

st.sidebar.button(
    "Reset Filter",
    on_click=reset_filter,
    use_container_width=True
)

st.sidebar.markdown("---")
st.sidebar.markdown("### Keterangan Warna")

for nomor, warna in WARNA.items():
    st.sidebar.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:9px;
                    margin-bottom:8px;">
            <div style="width:16px;height:16px;background:{warna};
                        border-radius:4px;"></div>
            <span>Klaster {nomor}</span>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# 10. FILTER DATA
# =========================================================

hasil_filter = hasil.copy()

kata = st.session_state["kata_cari"].strip()

if kata:
    hasil_filter = hasil_filter[
        hasil_filter["Kabupaten/Kota"].str.contains(
            kata, case=False, na=False
        )
    ]

hasil_filter = hasil_filter[
    hasil_filter["Kabupaten/Kota"].isin(
        st.session_state["wilayah_pilihan"]
    )
]

hasil_filter = hasil_filter[
    hasil_filter["Klaster"].isin(
        st.session_state["klaster_pilihan"]
    )
]


# =========================================================
# 11. NAVIGASI
# =========================================================

menu = [
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
    menu,
    index=0
)


# =========================================================
# 12. BERANDA
# =========================================================

if halaman == "Beranda":

    st.subheader("Ringkasan Dashboard")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        tampilkan_kartu(
            "Jumlah Wilayah",
            str(hasil["Kabupaten/Kota"].nunique())
        )

    with c2:
        tampilkan_kartu(
            "Jumlah Indikator",
            "9"
        )

    with c3:
        tampilkan_kartu(
            "K Terbaik",
            str(K_OPTIMAL)
        )

    with c4:
        tampilkan_kartu(
            "DBI Terendah",
            f"{DBI_TERBAIK:.4f}"
        )

    st.markdown("### Tentang Dashboard")

    st.write(
        "Dashboard ini menampilkan hasil klasterisasi kabupaten/kota "
        "di Provinsi Sumatera Barat berdasarkan sembilan indikator "
        "sosial ekonomi menggunakan algoritma K-Means."
    )

    st.write(
        "Nilai K terbaik ditentukan melalui evaluasi Davies-Bouldin "
        "Index (DBI). Nilai DBI terendah pada tabel evaluasi digunakan "
        "sebagai dasar pemilihan jumlah klaster."
    )

    st.info(
        f"Hasil evaluasi pada file Excel menunjukkan K terbaik = "
        f"{K_OPTIMAL}, dengan DBI = {DBI_TERBAIK:.4f}."
    )

    st.markdown("### Ringkasan Anggota Klaster")

    ringkasan = (
        hasil.groupby("Nama Klaster")["Kabupaten/Kota"]
        .agg(
            **{
                "Jumlah Wilayah": "nunique",
                "Anggota Wilayah": lambda x: ", ".join(
                    sorted(x.astype(str))
                )
            }
        )
       
