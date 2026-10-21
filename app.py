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
# 1. KONFIGURASI DASHBOARD
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

WARNA_KLASTER = {
    "Klaster 1": "#2E8B57",
    "Klaster 2": "#F2C94C",
    "Klaster 3": "#E74C3C",
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

    .label-kartu {
        color: #536579;
        font-size: 14px;
    }

    .nilai-kartu {
        color: #174A7E;
        font-size: 27px;
        font-weight: bold;
        margin-top: 8px;
    }

    .keterangan {
        color: #536579;
        font-size: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 3. FUNGSI BANTUAN
# =========================================================

def normalisasi_teks(teks):
    """Menyamakan format teks untuk pencocokan kolom dan wilayah."""
    if pd.isna(teks):
        return ""

    teks = str(teks).strip().lower()
    teks = unicodedata.normalize("NFKD", teks)
    teks = "".join(
        karakter
        for karakter in teks
        if not unicodedata.combining(karakter)
    )
    teks = teks.replace("_", " ")
    teks = re.sub(r"[^a-z0-9 ]", " ", teks)
    teks = re.sub(r"\s+", " ", teks).strip()

    return teks


def normalisasi_wilayah(nama):
    """Menyamakan penulisan nama kabupaten/kota."""
    nama = normalisasi_teks(nama)

    nama = nama.replace("kota sawah lunto", "kota sawahlunto")
    nama = nama.replace("sawah lunto", "sawahlunto")

    for awalan in ("kabupaten ", "kab ", "kab. "):
        if nama.startswith(awalan):
            nama = nama[len(awalan):].strip()
            break

    return nama


def cari_kolom(data, daftar_nama):
    """Mencari kolom berdasarkan beberapa kemungkinan nama."""
    if data is None or data.empty:
        return None

    kolom_asli = list(data.columns)
    kolom_normal = {
        normalisasi_teks(kolom): kolom
        for kolom in kolom_asli
    }

    # Prioritas pertama: nama yang cocok persis.
    for nama in daftar_nama:
        kunci = normalisasi_teks(nama)
        if kunci in kolom_normal:
            return kolom_normal[kunci]

    # Prioritas kedua: nama kolom yang mengandung kata kunci.
    for nama in daftar_nama:
        kunci = normalisasi_teks(nama)
        if len(kunci) < 4:
            continue

        for kolom in kolom_asli:
            if kunci in normalisasi_teks(kolom):
                return kolom

    return None


def ubah_ke_angka(kolom):
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


def nama_klaster(nomor):
    if pd.isna(nomor):
        return "Tidak diketahui"

    return f"Klaster {int(nomor)}"


def format_tabel(data):
    """Menampilkan tabel dengan lebar mengikuti area dashboard."""
    st.dataframe(
        data,
        hide_index=True,
        use_container_width=True
    )


# =========================================================
# 4. MEMBACA DAN MENYIAPKAN HASIL KLASTER
# =========================================================

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
            "Daerah",
        ]
    )

    kolom_klaster = cari_kolom(
        data,
        [
            "Klaster",
            "Cluster",
            "Cluster Label",
            "Label Klaster",
            "Hasil Klaster",
        ]
    )

    if kolom_wilayah is None:
        raise ValueError(
            "Kolom nama kabupaten/kota tidak ditemukan pada sheet "
            f"{SHEET_HASIL}. Kolom yang terbaca: "
            + ", ".join(map(str, data.columns))
        )

    if kolom_klaster is None:
        raise ValueError(
            "Kolom klaster tidak ditemukan pada sheet "
            f"{SHEET_HASIL}. Kolom yang terbaca: "
            + ", ".join(map(str, data.columns))
        )

    data = data.rename(
        columns={
            kolom_wilayah: "Kabupaten/Kota",
            kolom_klaster: "Klaster",
        }
    )

    data["Kabupaten/Kota"] = (
        data["Kabupaten/Kota"].astype(str).str.strip()
    )

    # Mendukung label seperti 1, Cluster 1, atau Klaster 1.
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
    ).copy()

    data["Klaster"] = data["Klaster"].astype(int)
    data["_kunci"] = data["Kabupaten/Kota"].apply(
        normalisasi_wilayah
    )
    data["Nama Klaster"] = data["Klaster"].apply(nama_klaster)

    return data


# =========================================================
# 5. MEMBACA FILE EXCEL
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
# 6. MENYIAPKAN EVALUASI DBI
# =========================================================

def siapkan_dbi(data):
    data = data.copy()

    kolom_k = cari_kolom(
        data,
        [
            "K",
            "Jumlah Klaster",
            "Jumlah_Klaster",
            "Jumlah Cluster",
            "Jumlah Cluster (K)",
        ]
    )

    kolom_dbi = cari_kolom(
        data,
        [
            "DBI",
            "Nilai DBI",
            "DBI Score",
            "Davies Bouldin Index",
            "Davies-Bouldin Index",
        ]
    )

    if kolom_k is None or kolom_dbi is None:
        raise ValueError(
            "Kolom K atau DBI tidak ditemukan pada sheet "
            f"{SHEET_DBI}. Kolom yang terbaca: "
            + ", ".join(map(str, data.columns))
        )

    data = data.rename(
        columns={
            kolom_k: "K",
            kolom_dbi: "DBI",
        }
    )

    data["K"] = ubah_ke_angka(data["K"])
    data["DBI"] = ubah_ke_angka(data["DBI"])

    data = data.dropna(subset=["K", "DBI"]).copy()
    data = data[data["K"] >= 2].copy()
    data["K"] = data["K"].astype(int)
    data = data.sort_values("K").reset_index(drop=True)

    if data.empty:
        raise ValueError(
            "Tidak ada nilai K dan DBI numerik yang dapat digunakan."
        )

    return data


# =========================================================
# 7. MEMBACA SHAPEFILE DARI ZIP
# =========================================================

@st.cache_data
def baca_shapefile(nama_zip):
    folder = tempfile.mkdtemp(prefix="peta_sumbar_")

    with zipfile.ZipFile(nama_zip, "r") as arsip:
        arsip.extractall(folder)

    daftar_shp = []

    for akar, _, berkas in os.walk(folder):
        for nama in berkas:
            if nama.lower().endswith(".shp"):
                daftar_shp.append(os.path.join(akar, nama))

    if not daftar_shp:
        raise ValueError(
            "File ZIP tidak berisi file shapefile (.shp)."
        )

    kandidat = [
        lokasi
        for lokasi in daftar_shp
        if "kab_kota" in os.path.basename(lokasi).lower()
    ]

    lokasi_shp = kandidat[0] if kandidat else daftar_shp[0]

    peta = gpd.read_file(lokasi_shp)

    if peta.empty:
        raise ValueError("Shapefile berhasil dibaca, tetapi tidak berisi data.")

    if peta.crs is not None and peta.crs.to_epsg() != 4326:
        peta = peta.to_crs(epsg=4326)

    kolom_nama = cari_kolom(
        peta,
        [
            "kab_kota",
            "Kabupaten/Kota",
            "WADMKK",
            "NAMOBJ",
            "Nama",
        ]
    )

    if kolom_nama is None:
        kolom_teks = [
            kolom
            for kolom in peta.columns
            if kolom != peta.geometry.name
            and peta[kolom].dtype == "object"
        ]

        if not kolom_teks:
            raise ValueError(
                "Kolom nama wilayah tidak ditemukan pada shapefile."
            )

        kolom_nama = kolom_teks[0]

    peta = peta.rename(
        columns={kolom_nama: "Nama_Peta"}
    )

    peta["Nama_Peta"] = peta["Nama_Peta"].astype(str).str.strip()
    peta["_kunci"] = peta["Nama_Peta"].apply(
        normalisasi_wilayah
    )

    geojson = json.loads(peta.to_json())

    # Menambahkan kunci pencocokan ke properti setiap fitur GeoJSON.
    for fitur in geojson["features"]:
        properti = fitur.setdefault("properties", {})
        properti["_kunci"] = normalisasi_wilayah(
            properti.get("Nama_Peta", "")
        )

    return peta, geojson


# =========================================================
# 8. HEADER DASHBOARD
# =========================================================

kolom_gambar, kolom_header = st.columns(
    [1, 5],
    vertical_alignment="center"
)

with kolom_gambar:
    if os.path.exists("rumah_gadang.png"):
        st.image("rumah_gadang.png", width=145)
    elif os.path.exists("rumah_gadang.jpg"):
        st.image("rumah_gadang.jpg", width=145)
    else:
        st.markdown(
            """
            <div style="font-size:70px; text-align:center;">
                🏠
            </div>
            """,
            unsafe_allow_html=True
        )
        st.caption("Rumah Gadang")

with kolom_header:
    st.markdown(
        """
        <h1 style="margin-bottom:0;">Dashboard Sumbar</h1>
        <h3 style="margin-top:5px; color:#536579;">
            Klasterisasi Indikator Sosial Ekonomi
        </h3>
        <p style="color:#65758B;">
            Provinsi Sumatera Barat | Algoritma K-Means
        </p>
        """,
        unsafe_allow_html=True
    )

st.divider()


# =========================================================
# 9. MEMASTIKAN FILE TERSEDIA
# =========================================================

if not os.path.exists(FILE_EXCEL):
    st.error(
        f"File Excel '{FILE_EXCEL}' tidak ditemukan. "
        "Pastikan file tersebut sudah diunggah ke repositori GitHub "
        "dan nama file sesuai dengan konfigurasi."
    )
    st.stop()

try:
    hasil_asli, evaluasi_asli, profil, pca = baca_excel(FILE_EXCEL)
    hasil = siapkan_hasil(hasil_asli)
    evaluasi_dbi = siapkan_dbi(evaluasi_asli)

except Exception as error:
    st.error(f"Data gagal dibaca: {error}")
    st.info(
        "Periksa nama sheet dan nama kolom pada file Excel. "
        "Nama sheet yang digunakan adalah Hasil_Klaster, "
        "Evaluasi_DBI, Profil_Klaster, dan Koordinat_PCA."
    )
    st.stop()


# =========================================================
# 10. MENENTUKAN K TERBAIK DARI DBI TERENDAH
# =========================================================

baris_terbaik = evaluasi_dbi.loc[
    evaluasi_dbi["DBI"].idxmin()
]

K_OPTIMAL = int(baris_terbaik["K"])
DBI_TERBAIK = float(baris_terbaik["DBI"])


# =========================================================
# 11. SIDEBAR DAN FILTER
# =========================================================

st.sidebar.title("Navigasi Dashboard")

menu = [
    "Beranda",
    "Hasil Klaster",
    "Evaluasi DBI",
    "Distribusi Klaster",
    "Profil Klaster",
    "Visualisasi PCA",
    "Peta Klaster",
]

halaman = st.sidebar.radio(
    "Pilih halaman",
    menu,
    index=0
)

st.sidebar.divider()
st.sidebar.subheader("Filter Data")

daftar_wilayah = sorted(
    hasil["Kabupaten/Kota"].dropna().unique().tolist()
)

daftar_klaster = sorted(
    hasil["Klaster"].dropna().unique().tolist()
)

kata_cari = st.sidebar.text_input(
    "Cari kabupaten/kota",
    placeholder="Contoh: Agam"
)

wilayah_pilihan = st.sidebar.multiselect(
    "Pilih wilayah",
    options=daftar_wilayah,
    default=daftar_wilayah
)

klaster_pilihan = st.sidebar.multiselect(
    "Pilih klaster",
    options=daftar_klaster,
    default=daftar_klaster,
    format_func=lambda x: f"Klaster {x}"
)

if st.sidebar.button("Reset Filter", use_container_width=True):
    st.rerun()

st.sidebar.divider()
st.sidebar.subheader("Keterangan Warna")

for nomor, warna in [
    (1, "#2E8B57"),
    (2, "#F2C94C"),
    (3, "#E74C3C"),
]:
    st.sidebar.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:9px;
                    margin-bottom:8px;">
            <div style="width:16px; height:16px; background:{warna};
                        border-radius:4px;"></div>
            <span>Klaster {nomor}</span>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# 12. MENERAPKAN FILTER
# =========================================================

hasil_filter = hasil.copy()

if kata_cari.strip():
    hasil_filter = hasil_filter[
        hasil_filter["Kabupaten/Kota"].str.contains(
            kata_cari.strip(),
            case=False,
            na=False
        )
    ]

hasil_filter = hasil_filter[
    hasil_filter["Kabupaten/Kota"].isin(wilayah_pilihan)
]

hasil_filter = hasil_filter[
    hasil_filter["Klaster"].isin(klaster_pilihan)
]


# =========================================================
# 13. HALAMAN BERANDA
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
        "Dashboard ini menyajikan hasil klasterisasi 19 kabupaten/kota "
        "di Provinsi Sumatera Barat berdasarkan sembilan indikator "
        "sosial ekonomi menggunakan algoritma K-Means."
    )

    st.write(
        "Evaluasi Davies-Bouldin Index (DBI) digunakan untuk memilih "
        "jumlah klaster. Nilai DBI yang lebih rendah menunjukkan "
        "hasil klaster yang lebih baik berdasarkan indeks tersebut."
    )

    st.info(
        f"Berdasarkan data evaluasi pada file Excel, K terbaik adalah "
        f"{K_OPTIMAL} dengan nilai DBI {DBI_TERBAIK:.4f}."
    )

    st.markdown("### Jumlah Wilayah Setiap Klaster")

    ringkasan = (
        hasil.groupby("Nama Klaster")["Kabupaten/Kota"]
        .nunique()
        .reset_index(name="Jumlah Wilayah")
    )

    ringkasan["Warna"] = ringkasan["Nama Klaster"].map(
        WARNA_KLASTER
    )

    fig = px.bar(
        ringkasan,
        x="Nama Klaster",
        y="Jumlah Wilayah",
        color="Nama Klaster",
        color_discrete_map=WARNA_KLASTER,
        text="Jumlah Wilayah",
        title="Jumlah Kabupaten/Kota pada Setiap Klaster"
    )

    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False
    )

    st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 14. HALAMAN HASIL KLASTER
# =========================================================

elif halaman == "Hasil Klaster":

    st.subheader("Hasil Klasterisasi Kabupaten/Kota")

    st.write(
        "Halaman ini menampilkan kelompok wilayah yang dihasilkan "
        "oleh algoritma K-Means."
    )

    st.caption(
        f"Wilayah yang ditampilkan: {len(hasil_filter)} dari "
        f"{len(hasil)} wilayah."
    )

    kolom_tampil = [
        kolom for kolom in hasil_filter.columns
        if not kolom.startswith("_")
    ]

    format_tabel(
        hasil_filter[kolom_tampil].sort_values(
            ["Klaster", "Kabupaten/Kota"]
        )
    )

    st.markdown("### Anggota Setiap Klaster")

    for nomor in sorted(hasil["Klaster"].unique()):
        anggota = sorted(
            hasil.loc[
                hasil["Klaster"] == nomor,
                "Kabupaten/Kota"
            ].tolist()
        )

        with st.expander(f"Klaster {nomor} — {len(anggota)} wilayah"):
            if anggota:
                for wilayah in anggota:
                    st.write(f"• {wilayah}")
            else:
                st.write("Belum ada wilayah pada klaster ini.")


# =========================================================
# 15. HALAMAN EVALUASI DBI
# =========================================================

elif halaman == "Evaluasi DBI":

    st.subheader("Evaluasi Davies-Bouldin Index (DBI)")

    st.write(
        "Nilai DBI dibandingkan untuk setiap jumlah klaster yang "
        "diuji. Jumlah klaster dengan nilai DBI paling rendah "
        "dipilih sebagai K terbaik."
    )

    c1, c2 = st.columns(2)

    with c1:
        tampilkan_kartu(
            "K Terbaik",
            str(K_OPTIMAL)
        )

    with c2:
        tampilkan_kartu(
            "Nilai DBI Terendah",
            f"{DBI_TERBAIK:.4f}"
        )

    st.markdown("### Grafik Evaluasi DBI")

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=evaluasi_dbi["K"],
            y=evaluasi_dbi["DBI"],
            mode="lines+markers+text",
            text=evaluasi_dbi["DBI"].map(
                lambda nilai: f"{nilai:.4f}"
            ),
            textposition="top center",
            name="Nilai DBI",
            line=dict(color=WARNA_BIRU, width=3),
            marker=dict(size=9)
        )
    )

    fig.add_trace(
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

    fig.update_layout(
        xaxis_title="Jumlah Klaster (K)",
        yaxis_title="Davies-Bouldin Index",
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="x unified",
        xaxis=dict(
            tickmode="array",
            tickvals=evaluasi_dbi["K"].tolist(),
            dtick=1
        )
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Tabel Evaluasi DBI")

    tabel_dbi = evaluasi_dbi.copy()
    tabel_dbi["DBI"] = tabel_dbi["DBI"].round(4)
    tabel_dbi["Status"] = tabel_dbi["K"].apply(
        lambda k: "K Terbaik" if k == K_OPTIMAL else ""
    )

    format_tabel(tabel_dbi)

    st.success(
        f"K terbaik adalah {K_OPTIMAL}, dengan nilai DBI "
        f"{DBI_TERBAIK:.4f}."
    )


# =========================================================
# 16. HALAMAN DISTRIBUSI KLASTER
# =========================================================

elif halaman == "Distribusi Klaster":

    st.subheader("Distribusi Klaster")

    data_distribusi = (
        hasil_filter.groupby("Nama Klaster")["Kabupaten/Kota"]
        .nunique()
        .reset_index(name="Jumlah Wilayah")
    )

    if data_distribusi.empty:
        st.warning(
            "Tidak ada data yang sesuai dengan filter yang dipilih."
        )
    else:
        c1, c2 = st.columns(2)

        with c1:
            fig_pie = px.pie(
                data_distribusi,
                names="Nama Klaster",
                values="Jumlah Wilayah",
                color="Nama Klaster",
                color_discrete_map=WARNA_KLASTER,
                hole=0.35,
                title="Proporsi Wilayah Setiap Klaster"
            )

            fig_pie.update_traces(
                textposition="inside",
                textinfo="label+value+percent"
            )

            st.plotly_chart(
                fig_pie,
                use_container_width=True
            )

        with c2:
            fig_bar = px.bar(
                data_distribusi,
                x="Nama Klaster",
                y="Jumlah Wilayah",
                color="Nama Klaster",
                color_discrete_map=WARNA_KLASTER,
                text="Jumlah Wilayah",
                title="Jumlah Wilayah Setiap Klaster"
            )

            fig_bar.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                showlegend=False
            )

            st.plotly_chart(
                fig_bar,
                use_container_width=True
            )

        st.markdown("### Rincian Distribusi")

        format_tabel(data_distribusi)


# =========================================================
# 17. HALAMAN PROFIL KLASTER
# =========================================================

elif halaman == "Profil Klaster":

    st.subheader("Profil Klaster")

    if profil.empty:
        st.warning(
            f"Sheet '{SHEET_PROFIL}' tidak ditemukan atau kosong. "
            "Periksa kembali file Excel."
        )
    else:
        st.write(
            "Tabel berikut menampilkan profil klaster sebagaimana "
            "tersimpan pada sheet Profil_Klaster."
        )

        format_tabel(profil)

        # Grafik hanya dibuat jika ditemukan kolom klaster
        # dan terdapat kolom indikator numerik.
        kolom_klaster_profil = cari_kolom(
            profil,
            ["Klaster", "Cluster", "Nama Klaster"]
        )

        kolom_indikator = [
            kolom for kolom in profil.columns
            if kolom != kolom_klaster_profil
            and pd.api.types.is_numeric_dtype(profil[kolom])
        ]

        if kolom_klaster_profil is not None and kolom_indikator:
            data_grafik = profil.copy()

            data_grafik[kolom_klaster_profil] = (
                data_grafik[kolom_klaster_profil]
                .astype(str)
            )

            pilihan_indikator = st.selectbox(
                "Pilih indikator untuk dibandingkan",
                kolom_indikator
            )

            fig = px.bar(
                data_grafik,
                x=kolom_klaster_profil,
                y=pilihan_indikator,
                color=kolom_klaster_profil,
                color_discrete_map=WARNA_KLASTER,
                text=pilihan_indikator,
                title=f"Perbandingan {pilihan_indikator} per Klaster"
            )

            fig.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                showlegend=False
            )

            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(
                "Grafik profil otomatis tidak ditampilkan karena "
                "kolom klaster atau indikator numerik belum dapat "
                "dikenali. Tabel profil tetap ditampilkan."
            )


# =========================================================
# 18. HALAMAN VISUALISASI PCA
# =========================================================

elif halaman == "Visualisasi PCA":

    st.subheader("Visualisasi Persebaran Wilayah dengan PCA")

    st.write(
        "PCA digunakan untuk memvisualisasikan persebaran titik "
        "wilayah dalam dua dimensi. PCA pada halaman ini berfungsi "
        "sebagai visualisasi, bukan sebagai metode penentuan K."
    )

    if pca.empty:
        st.warning(
            f"Sheet '{SHEET_PCA}' tidak ditemukan atau kosong. "
            "Periksa kembali file Excel."
        )
    else:
        kolom_wilayah_pca = cari_kolom(
            pca,
            [
                "Kabupaten/Kota",
                "Kabupaten Kota",
                "kab_kota",
                "Wilayah",
                "Daerah",
            ]
        )

        kolom_pc1 = cari_kolom(
            pca,
            ["PC1", "PCA1", "Komponen Utama 1", "Principal Component 1"]
        )

        kolom_pc2 = cari_kolom(
            pca,
            ["PC2", "PCA2", "Komponen Utama 2", "Principal Component 2"]
        )

        kolom_klaster_pca = cari_kolom(
            pca,
            ["Klaster", "Cluster", "Label Klaster"]
        )

        if (
            kolom_wilayah_pca is None
            or kolom_pc1 is None
            or kolom_pc2 is None
        ):
            st.warning(
                "Kolom wilayah, PC1, atau PC2 tidak ditemukan pada "
                f"sheet {SHEET_PCA}. Kolom yang tersedia: "
                + ", ".join(map(str, pca.columns))
            )
            format_tabel(pca)

        else:
            data_pca = pca.copy()

            data_pca = data_pca.rename(
                columns={
                    kolom_wilayah_pca: "Kabupaten/Kota",
                    kolom_pc1: "PC1",
                    kolom_pc2: "PC2",
                }
            )

            data_pca["PC1"] = ubah_ke_angka(data_pca["PC1"])
            data_pca["PC2"] = ubah_ke_angka(data_pca["PC2"])

            if kolom_klaster_pca is not None:
                data_pca = data_pca.rename(
                    columns={kolom_klaster_pca: "Klaster"}
                )
                data_pca["Klaster"] = (
                    data_pca["Klaster"].astype(str)
                    .str.extract(r"(\d+)", expand=False)
                )
            else:
                data_pca = data_pca.merge(
                    hasil[["Kabupaten/Kota", "Klaster"]],
                    on="Kabupaten/Kota",
                    how="left"
                )
                data_pca["Klaster"] = data_pca["Klaster"].astype(str)

            data_pca["Nama Klaster"] = (
                data_pca["Klaster"].apply(
                    lambda x: nama_klaster(
                        pd.to_numeric(x, errors="coerce")
                    )
                )
            )

            data_pca = data_pca.dropna(
                subset=["PC1", "PC2"]
            )

            if data_pca.empty:
                st.warning(
                    "Tidak ada koordinat PC1 dan PC2 yang valid."
                )
            else:
                data_pca = data_pca[
                    data_pca["Kabupaten/Kota"].isin(
                        hasil_filter["Kabupaten/Kota"]
                    )
                ]

                fig = px.scatter(
                    data_pca,
                    x="PC1",
                    y="PC2",
                    color="Nama Klaster",
                    color_discrete_map=WARNA_KLASTER,
                    hover_name="Kabupaten/Kota",
                    text="Kabupaten/Kota",
                    title="Persebaran Kabupaten/Kota Berdasarkan PCA"
                )

                fig.update_traces(
                    textposition="top center",
                    marker=dict(size=12)
                )

                fig.update_layout(
                    xaxis_title="Komponen Utama 1 (PC1)",
                    yaxis_title="Komponen Utama 2 (PC2)",
                    plot_bgcolor="white",
                    paper_bgcolor="white"
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )


# =========================================================
# 19. HALAMAN PETA KLASTER
# =========================================================

elif halaman == "Peta Klaster":

    st.subheader("Peta Klaster Kabupaten/Kota Sumatera Barat")

    st.write(
        "Warna pada peta menunjukkan klaster setiap kabupaten/kota "
        "berdasarkan hasil K-Means."
    )

    if not os.path.exists(FILE_ZIP):
        st.error(
            f"File peta '{FILE_ZIP}' tidak ditemukan. "
            "Pastikan ZIP shapefile sudah diunggah ke repositori GitHub."
        )
        st.stop()

    try:
        peta, geojson = baca_shapefile(FILE_ZIP)

        # Merge menggunakan kunci nama wilayah yang sudah dinormalisasi.
        data_peta = peta.merge(
            hasil[
                [
                    "Kabupaten/Kota",
                    "Klaster",
                    "Nama Klaster",
                    "_kunci",
                ]
            ],
            on="_kunci",
            how="left",
            suffixes=("", "_hasil")
        )

        # Memastikan klaster tetap numerik setelah proses merge.
        data_peta["Klaster"] = pd.to_numeric(
            data_peta["Klaster"],
            errors="coerce"
        )

        data_peta["Nama Klaster"] = data_peta["Klaster"].apply(
            lambda x: nama_klaster(x) if pd.notna(x) else "Belum cocok"
        )

        data_peta["Nama Wilayah"] = data_peta[
            "Kabupaten/Kota"
        ].fillna(data_peta["Nama_Peta"])

        # GeoJSON dan dataframe harus memakai kunci yang sama.
        data_peta_json = json.loads(data_peta.to_json())

        for fitur in data_peta_json["features"]:
            properti = fitur.setdefault("properties", {})
            properti["_kunci"] = normalisasi_wilayah(
                properti.get("Nama_Peta", "")
            )

        # Filter peta berdasarkan pilihan pengguna.
        kunci_wilayah_terpilih = set(
            hasil_filter["_kunci"].tolist()
        )

        if kata_cari.strip():
            data_peta = data_peta[
                data_peta["_kunci"].isin(kunci_wilayah_terpilih)
            ]

        if len(klaster_pilihan) < len(daftar_klaster):
            data_peta = data_peta[
                data_peta["Klaster"].isin(klaster_pilihan)
            ]

        if data_peta.empty:
            st.warning(
                "Tidak ada wilayah pada peta yang sesuai dengan filter."
            )
        else:
            # Hanya fitur yang ditampilkan dimasukkan ke GeoJSON.
            data_peta_json = json.loads(data_peta.to_json())

            for fitur in data_peta_json["features"]:
                properti = fitur.setdefault("properties", {})
                properti["_kunci"] = normalisasi_wilayah(
                    properti.get("Nama_Peta", "")
                )

            fig = px.choropleth_mapbox(
                data_peta,
                geojson=data_peta_json,
                locations="_kunci",
                featureidkey="properties._kunci",
                color="Nama Klaster",
                color_discrete_map=WARNA_KLASTER,
                category_orders={
                    "Nama Klaster": [
                        "Klaster 1",
                        "Klaster 2",
                        "Klaster 3",
                        "Belum cocok",
                    ]
                },
                hover_name="Nama Wilayah",
                hover_data={
                    "Klaster": True,
                    "_kunci": False,
                    "Nama Klaster": True,
                },
                mapbox_style="carto-positron",
                zoom=5.8,
                center={
                    "lat": -0.65,
                    "lon": 100.65
                },
                opacity=0.65,
                title="Peta Persebaran Klaster Sumatera Barat"
            )

            fig.update_traces(
                marker_line_width=1.2,
                marker_line_color="#FFFFFF"
            )

            fig.update_layout(
                margin={"r": 0, "t": 50, "l": 0, "b": 0},
                legend_title_text="Hasil Klaster",
                height=650
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

            st.markdown("### Keterangan Wilayah")

            tabel_peta = data_peta[
                [
                    "Nama_Wilayah"
                    if "Nama_Wilayah" in data_peta.columns
                    else "Nama_Peta",
                    "Klaster",
                    "Nama Klaster",
                ]
            ].copy()

            tabel_peta.columns = [
                "Kabupaten/Kota",
                "Nomor Klaster",
                "Nama Klaster",
            ]

            tabel_peta = tabel_peta.sort_values(
                "Kabupaten/Kota"
            )

            format_tabel(tabel_peta)

            # Memperlihatkan wilayah yang belum cocok agar dapat diperbaiki.
            wilayah_tidak_cocok = data_peta[
                data_peta["Klaster"].isna()
            ]["Nama_Peta"].tolist()

            if wilayah_tidak_cocok:
                st.warning(
                    "Beberapa nama wilayah pada shapefile belum cocok "
                    "dengan data hasil klaster: "
                    + ", ".join(wilayah_tidak_cocok)
                )

    except Exception as error:
        st.error(f"Peta gagal ditampilkan: {error}")
        st.info(
            "Pastikan ZIP berisi shapefile lengkap (.shp, .shx, .dbf, "
            "dan .prj jika tersedia), nama kolom wilayah sesuai, dan "
            "dependensi pada requirements.txt sudah terpasang."
        )


# =========================================================
# 20. FOOTER
# =========================================================

st.divider()

st.caption(
    "Dashboard Klasterisasi Indikator Sosial Ekonomi "
    "Provinsi Sumatera Barat | Metode K-Means dan evaluasi DBI"
)
