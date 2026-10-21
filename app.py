
import json
import re
import zipfile
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Geopandas digunakan untuk membaca shapefile peta.
try:
    import geopandas as gpd
except ImportError:
    gpd = None


# =========================================================
# 1. KONFIGURASI DASHBOARD
# =========================================================

st.set_page_config(
    page_title="Dashboard Klasterisasi Sumatera Barat",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_SHAPEFILE = "Provinsi Sumatera Barat-KAB_KOTA.zip"
FILE_RUMAH_GADANG = "rumah_gadang.png"

WARNA_KLASTER = {
    "Klaster 1": "#2E8B57",
    "Klaster 2": "#F6C85F",
    "Klaster 3": "#D9534F",
    "Belum cocok": "#9E9E9E",
}

URUTAN_KLASTER = [
    "Klaster 1",
    "Klaster 2",
    "Klaster 3",
    "Belum cocok",
]

INDIKATOR = [
    "Tingkat_Kemiskinan (%)",
    "PDRB_per_Kapita (Ribu Rp)",
    "TPT (%)",
    "Gini_Ratio",
    "Umur Harapan Hidup(Tahun)",
    "Harapan Lama Sekolah(Tahun)",
    "Rata-rata Lama Sekolah(Tahun)",
    "Pengeluaran Per Kapita(Ribu Rupiah/Orang/Tahun)",
    "Laju Pertumbuhan Ekonomi",
]


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
        background-color: #EAF2FB;
    }

    .judul-dashboard {
        color: #123D69;
        font-size: 30px;
        font-weight: 800;
        margin-bottom: 2px;
    }

    .subjudul-dashboard {
        color: #4A6683;
        font-size: 16px;
        margin-top: 0;
    }

    .kartu-info {
        background-color: white;
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #DCE6F1;
    }

    h1, h2, h3 {
        color: #123D69;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 3. FUNGSI BANTU
# =========================================================

def kunci_teks(teks):
    """Menyeragamkan teks agar nama wilayah mudah dicocokkan."""
    teks = str(teks).strip().lower()
    teks = teks.replace("&", "dan")
    teks = re.sub(r"^(kabupaten|kab\.|kota)\s+", "", teks)
    teks = re.sub(r"[^a-z0-9]", "", teks)
    return teks


def cari_kolom(data, kandidat):
    """Mencari nama kolom yang sesuai meskipun berbeda kapital/spasi."""
    if data is None or data.empty:
        return None

    peta = {
        str(kolom).strip().lower(): kolom
        for kolom in data.columns
    }

    for nama in kandidat:
        if nama.lower() in peta:
            return peta[nama.lower()]

    # Pencocokan sebagian untuk mengakomodasi nama kolom berbeda.
    for kolom in data.columns:
        nama_kolom = str(kolom).strip().lower()
        for nama in kandidat:
            if nama.lower() in nama_kolom:
                return kolom

    return None


def kolom_numerik(data, kolom):
    """Mengubah sebuah kolom menjadi angka bila memungkinkan."""
    if kolom is None or kolom not in data.columns:
        return pd.Series(dtype=float)

    return pd.to_numeric(data[kolom], errors="coerce")


def nama_klaster(nilai):
    """Mengubah nilai klaster menjadi label yang konsisten."""
    if pd.isna(nilai):
        return "Belum cocok"

    teks = str(nilai).strip()
    angka = re.search(r"\d+", teks)

    if angka:
        nomor = int(angka.group())
        if nomor in (1, 2, 3):
            return f"Klaster {nomor}"

    return "Belum cocok"


def cari_sheet(semua_sheet, kandidat):
    """Memilih sheet berdasarkan nama, dengan fallback jika tidak ditemukan."""
    for nama in kandidat:
        for nama_asli, data in semua_sheet.items():
            if nama.lower() == nama_asli.strip().lower():
                return data.copy()

    for nama_asli, data in semua_sheet.items():
        nama_normal = nama_asli.strip().lower().replace(" ", "_")
        for nama in kandidat:
            if nama.lower() in nama_normal:
                return data.copy()

    return pd.DataFrame()


def tampilkan_tabel(data, tinggi=420):
    """Menampilkan tabel dengan lebar kolom menyesuaikan tampilan."""
    st.dataframe(
        data,
        use_container_width=True,
        height=tinggi,
        hide_index=True,
    )


# =========================================================
# 4. MEMBACA FILE EXCEL
# =========================================================

@st.cache_data
def baca_excel(nama_file):
    semua_sheet = pd.read_excel(
        nama_file,
        sheet_name=None,
        engine="openpyxl",
    )

    hasil = cari_sheet(
        semua_sheet,
        ["Hasil_Klaster", "Hasil Klaster", "hasil_cluster"],
    )

    evaluasi = cari_sheet(
        semua_sheet,
        ["Evaluasi_DBI", "Evaluasi DBI", "evaluasi_dbi"],
    )

    profil = cari_sheet(
        semua_sheet,
        ["Profil_Klaster", "Profil Klaster", "profil_cluster"],
    )

    pca = cari_sheet(
        semua_sheet,
        ["Koordinat_PCA", "Koordinat PCA", "PCA", "Visualisasi_PCA"],
    )

    # Jika nama sheet tidak cocok, gunakan sheet pertama sebagai hasil.
    if hasil.empty and semua_sheet:
        hasil = next(iter(semua_sheet.values())).copy()

    return hasil, evaluasi, profil, pca, list(semua_sheet.keys())


# =========================================================
# 5. MEMUAT DATA
# =========================================================

if not Path(FILE_EXCEL).exists():
    st.error(
        f"File Excel '{FILE_EXCEL}' tidak ditemukan. "
        "Pastikan file tersebut sudah diunggah ke repositori GitHub "
        "pada folder yang sama dengan app.py."
    )
    st.stop()

try:
    data_asli, evaluasi_dbi, profil_klaster, data_pca, nama_sheet = (
        baca_excel(FILE_EXCEL)
    )
except Exception as error:
    st.error(f"File Excel gagal dibaca: {error}")
    st.stop()

if data_asli.empty:
    st.error(
        "Data hasil klaster tidak ditemukan. Periksa nama sheet "
        "dan isi file Excel."
    )
    st.stop()


# =========================================================
# 6. MENYERAGAMKAN DATA HASIL KLASTER
# =========================================================

kolom_wilayah = cari_kolom(
    data_asli,
    ["Kabupaten/Kota", "Kabupaten Kota", "Nama Wilayah", "Wilayah"],
)

kolom_klaster = cari_kolom(
    data_asli,
    ["Klaster", "Cluster", "Label Klaster", "Hasil Klaster"],
)

if kolom_wilayah is None:
    st.error(
        "Kolom nama wilayah tidak ditemukan pada sheet hasil klaster. "
        "Pastikan terdapat kolom Kabupaten/Kota."
    )
    st.write("Kolom yang terbaca:", list(data_asli.columns))
    st.stop()

if kolom_klaster is None:
    st.error(
        "Kolom klaster tidak ditemukan pada sheet hasil klaster. "
        "Pastikan terdapat kolom Klaster atau Cluster."
    )
    st.write("Kolom yang terbaca:", list(data_asli.columns))
    st.stop()

hasil = data_asli.copy()

hasil["Kabupaten/Kota"] = hasil[kolom_wilayah].astype(str).str.strip()
hasil["Klaster"] = hasil[kolom_klaster].apply(nama_klaster)
hasil["_kunci"] = hasil["Kabupaten/Kota"].apply(kunci_teks)

# Hanya mempertahankan wilayah yang memiliki nama.
hasil = hasil[
    hasil["Kabupaten/Kota"].notna()
    & (hasil["Kabupaten/Kota"].str.strip() != "")
].copy()

# Menambahkan kolom indikator jika tersedia pada Excel.
indikator_tersedia = []
for indikator in INDIKATOR:
    kolom = cari_kolom(hasil, [indikator])
    if kolom is not None and kolom not in indikator_tersedia:
        indikator_tersedia.append(kolom)

# Menghapus duplikasi wilayah bila ada, agar tabel hasil tidak berulang.
hasil = hasil.drop_duplicates(
    subset=["_kunci"],
    keep="first",
).reset_index(drop=True)

jumlah_wilayah = hasil["Kabupaten/Kota"].nunique()
jumlah_indikator = len(indikator_tersedia)

# =========================================================
# 7. MEMBACA HASIL EVALUASI DBI
# =========================================================

kolom_k = cari_kolom(
    evaluasi_dbi,
    ["K", "Jumlah Klaster", "Jumlah_Klaster"],
)

kolom_dbi = cari_kolom(
    evaluasi_dbi,
    ["DBI", "Davies Bouldin Index", "Davies-Bouldin Index"],
)

if not evaluasi_dbi.empty and kolom_k and kolom_dbi:
    evaluasi_dbi = evaluasi_dbi.copy()
    evaluasi_dbi["K"] = pd.to_numeric(
        evaluasi_dbi[kolom_k],
        errors="coerce",
    )
    evaluasi_dbi["DBI"] = pd.to_numeric(
        evaluasi_dbi[kolom_dbi],
        errors="coerce",
    )
    evaluasi_dbi = evaluasi_dbi.dropna(
        subset=["K", "DBI"]
    ).sort_values("K")
else:
    evaluasi_dbi = pd.DataFrame(columns=["K", "DBI"])

if not evaluasi_dbi.empty:
    baris_terbaik = evaluasi_dbi.loc[evaluasi_dbi["DBI"].idxmin()]
    k_optimal = int(baris_terbaik["K"])
    dbi_terbaik = float(baris_terbaik["DBI"])
else:
    k_optimal = hasil["Klaster"].nunique()
    dbi_terbaik = None


# =========================================================
# 8. JUDUL DASHBOARD
# =========================================================

kolom_gambar, kolom_judul = st.columns([1, 5])

with kolom_gambar:
    if Path(FILE_RUMAH_GADANG).exists():
        st.image(FILE_RUMAH_GADANG, use_container_width=True)
    else:
        st.markdown(
            """
            <div style="
                background-color:#DDEBFA;
                border-radius:14px;
                padding:18px;
                text-align:center;
                font-size:48px;">
                🏠
            </div>
            """,
            unsafe_allow_html=True,
        )

with kolom_judul:
    st.markdown(
        '<div class="judul-dashboard">Dashboard Sumbar</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subjudul-dashboard">'
        'Klasterisasi Indikator Sosial Ekonomi '
        'di Provinsi Sumatera Barat menggunakan K-Means'
        '</div>',
        unsafe_allow_html=True,
    )

st.divider()


# =========================================================
# 9. MENU NAVIGASI
# =========================================================

with st.sidebar:
    st.title("📊 Navigasi")
    st.caption("Dashboard Klasterisasi Sumatera Barat")

    menu = st.radio(
        "Pilih halaman:",
        [
            "Beranda",
            "Hasil Klaster",
            "Evaluasi DBI",
            "Distribusi Klaster",
            "Profil Klaster",
            "Visualisasi PCA",
            "Peta Klaster",
        ],
    )

    st.divider()
    st.caption("Sumber data: hasil pengolahan indikator sosial ekonomi.")


# =========================================================
# 10. HALAMAN BERANDA
# =========================================================

if menu == "Beranda":
    st.title("Beranda")
    st.write(
        "Dashboard ini menyajikan hasil klasterisasi kabupaten/kota "
        "di Provinsi Sumatera Barat berdasarkan indikator sosial ekonomi "
        "menggunakan algoritma K-Means."
    )

    metrik1, metrik2, metrik3 = st.columns(3)

    metrik1.metric(
        "Jumlah Wilayah",
        jumlah_wilayah,
    )

    metrik2.metric(
        "Jumlah Indikator",
        jumlah_indikator,
    )

    if dbi_terbaik is not None:
        metrik3.metric(
            "K Terbaik berdasarkan DBI",
            f"K = {k_optimal}",
            help="Nilai K dipilih berdasarkan nilai DBI terendah.",
        )
    else:
        metrik3.metric(
            "Jumlah Klaster pada Data",
            hasil["Klaster"].nunique(),
        )

    st.subheader("Tentang Dashboard")
    st.markdown(
        """
        Dashboard ini menyediakan beberapa informasi utama:

        - **Hasil Klaster:** melihat hasil pengelompokan setiap wilayah.
        - **Evaluasi DBI:** membandingkan nilai Davies-Bouldin Index
          untuk menentukan jumlah klaster yang lebih baik.
        - **Distribusi Klaster:** melihat jumlah dan persentase wilayah
          pada setiap klaster.
        - **Profil Klaster:** melihat karakteristik indikator pada
          setiap kelompok.
        - **Visualisasi PCA:** melihat persebaran titik wilayah dalam
          dua komponen utama.
        - **Peta Klaster:** melihat persebaran klaster secara geografis.
        """
    )

    st.subheader("Ringkasan Anggota Klaster")

    ringkasan = (
        hasil.groupby("Klaster")["Kabupaten/Kota"]
        .nunique()
        .reindex(URUTAN_KLASTER, fill_value=0)
        .reset_index()
    )
    ringkasan.columns = ["Klaster", "Jumlah Wilayah"]
    ringkasan = ringkasan[ringkasan["Jumlah Wilayah"] > 0]

    if not ringkasan.empty:
        fig = px.bar(
            ringkasan,
            x="Klaster",
            y="Jumlah Wilayah",
            color="Klaster",
            color_discrete_map=WARNA_KLASTER,
            category_orders={"Klaster": URUTAN_KLASTER},
            text="Jumlah Wilayah",
            title="Jumlah Wilayah pada Setiap Klaster",
        )
        fig.update_layout(
            showlegend=False,
            xaxis_title="Klaster",
            yaxis_title="Jumlah Wilayah",
        )
        st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 11. HALAMAN HASIL KLASTER
# =========================================================

elif menu == "Hasil Klaster":
    st.title("Hasil Klaster")
    st.write(
        "Halaman ini menampilkan hasil pengelompokan kabupaten/kota "
        "berdasarkan kemiripan sembilan indikator sosial ekonomi "
        "yang digunakan dalam penelitian."
    )

    kolom_filter1, kolom_filter2 = st.columns([2, 1])

    with kolom_filter1:
        pencarian = st.text_input(
            "Cari nama kabupaten/kota",
            placeholder="Contoh: Agam atau Kota Padang",
        )

    with kolom_filter2:
        pilihan_klaster = st.selectbox(
            "Filter klaster",
            ["Semua Klaster"] + [
                k for k in URUTAN_KLASTER
                if k in hasil["Klaster"].unique()
            ],
        )

    tampil = hasil.copy()

    if pencarian.strip():
        tampil = tampil[
            tampil["Kabupaten/Kota"]
            .str.contains(pencarian.strip(), case=False, na=False)
        ]

    if pilihan_klaster != "Semua Klaster":
        tampil = tampil[tampil["Klaster"] == pilihan_klaster]

    kolom_tampil = ["Kabupaten/Kota", "Klaster"]

    for kolom in indikator_tersedia:
        if kolom not in kolom_tampil:
            kolom_tampil.append(kolom)

    tampil = tampil[kolom_tampil].copy()

    st.write(f"Jumlah wilayah yang ditampilkan: **{len(tampil)}**")
    tampilkan_tabel(tampil)

    file_csv = tampil.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "Unduh hasil yang ditampilkan (CSV)",
        data=file_csv,
        file_name="hasil_klaster_sumatera_barat.csv",
        mime="text/csv",
    )


# =========================================================
# 12. HALAMAN EVALUASI DBI
# =========================================================

elif menu == "Evaluasi DBI":
    st.title("Evaluasi Davies-Bouldin Index (DBI)")

    st.write(
        "Davies-Bouldin Index digunakan untuk mengevaluasi kualitas "
        "hasil klasterisasi. Secara umum, nilai DBI yang lebih rendah "
        "menunjukkan hasil klaster yang lebih baik."
    )

    if evaluasi_dbi.empty:
        st.warning(
            "Data evaluasi DBI tidak ditemukan atau kolom K dan DBI "
            "belum terbaca. Pastikan sheet Evaluasi_DBI berisi kolom K "
            "dan DBI."
        )
    else:
        metrik1, metrik2 = st.columns(2)

        metrik1.metric(
            "K Terbaik",
            f"K = {k_optimal}",
        )

        metrik2.metric(
            "Nilai DBI Terendah",
            f"{dbi_terbaik:.4f}",
        )

        fig = px.line(
            evaluasi_dbi,
            x="K",
            y="DBI",
            markers=True,
            title="Perbandingan Nilai DBI pada Setiap K",
        )

        fig.update_traces(
            line=dict(width=3),
            marker=dict(size=9),
        )

        fig.update_layout(
            xaxis_title="Jumlah Klaster (K)",
            yaxis_title="Nilai DBI",
        )

        fig.update_xaxes(dtick=1)

        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Tabel Evaluasi DBI")
        tampilkan_tabel(evaluasi_dbi[["K", "DBI"]])

        st.info(
            f"Berdasarkan data evaluasi yang tersedia, nilai DBI terendah "
            f"adalah {dbi_terbaik:.4f} pada K = {k_optimal}. "
            "Nilai ini menjadi dasar pemilihan jumlah klaster."
        )


# =========================================================
# 13. HALAMAN DISTRIBUSI KLASTER
# =========================================================

elif menu == "Distribusi Klaster":
    st.title("Distribusi Klaster")

    jumlah_per_klaster = (
        hasil.groupby("Klaster")["Kabupaten/Kota"]
        .nunique()
        .reindex(URUTAN_KLASTER, fill_value=0)
        .reset_index()
    )

    jumlah_per_klaster.columns = ["Klaster", "Jumlah Wilayah"]
    jumlah_per_klaster = jumlah_per_klaster[
        jumlah_per_klaster["Jumlah Wilayah"] > 0
    ].copy()

    jumlah_per_klaster["Persentase (%)"] = (
        jumlah_per_klaster["Jumlah Wilayah"] / jumlah_wilayah * 100
    ).round(2)

    kolom1, kolom2 = st.columns(2)

    with kolom1:
        fig_bar = px.bar(
            jumlah_per_klaster,
            x="Klaster",
            y="Jumlah Wilayah",
            color="Klaster",
            color_discrete_map=WARNA_KLASTER,
            category_orders={"Klaster": URUTAN_KLASTER},
            text="Jumlah Wilayah",
            title="Jumlah Wilayah per Klaster",
        )
        fig_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

    with kolom2:
        fig_pie = px.pie(
            jumlah_per_klaster,
            names="Klaster",
            values="Jumlah Wilayah",
            color="Klaster",
            color_discrete_map=WARNA_KLASTER,
            category_orders={"Klaster": URUTAN_KLASTER},
            hole=0.35,
            title="Persentase Wilayah per Klaster",
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    st.subheader("Ringkasan Distribusi")
    tampilkan_tabel(jumlah_per_klaster)


# =========================================================
# 14. HALAMAN PROFIL KLASTER
# =========================================================

elif menu == "Profil Klaster":
    st.title("Profil Klaster")

    st.write(
        "Halaman ini menampilkan profil indikator sosial ekonomi "
        "untuk membantu membandingkan karakteristik setiap klaster."
    )

    if not profil_klaster.empty:
        st.subheader("Tabel Profil Klaster")
        tampilkan_tabel(profil_klaster)

        kolom_profil_klaster = cari_kolom(
            profil_klaster,
            ["Klaster", "Cluster", "Label Klaster"],
        )

        if kolom_profil_klaster:
            profil_plot = profil_klaster.copy()
            profil_plot["Klaster"] = (
                profil_plot[kolom_profil_klaster].apply(nama_klaster)
            )

            kolom_numerik_profil = [
                kolom for kolom in profil_plot.columns
                if kolom != kolom_profil_klaster
                and pd.api.types.is_numeric_dtype(profil_plot[kolom])
            ]

            if kolom_numerik_profil:
                indikator_pilihan = st.selectbox(
                    "Pilih indikator untuk dibandingkan",
                    kolom_numerik_profil,
                )

                fig = px.bar(
                    profil_plot,
                    x="Klaster",
                    y=indikator_pilihan,
                    color="Klaster",
                    color_discrete_map=WARNA_KLASTER,
                    category_orders={"Klaster": URUTAN_KLASTER},
                    title=f"Perbandingan {indikator_pilihan} per Klaster",
                )
                st.plotly_chart(fig, use_container_width=True)
    else:
        # Jika sheet profil tidak tersedia, coba hitung dari data hasil.
        if not indikator_tersedia:
            st.warning(
                "Sheet Profil_Klaster tidak ditemukan dan kolom indikator "
                "tidak tersedia pada sheet hasil klaster."
            )
        else:
            st.info(
                "Sheet Profil_Klaster tidak ditemukan. Profil dihitung "
                "dari nilai rata-rata indikator yang tersedia pada data hasil."
            )

            data_profil = hasil.copy()

            for kolom in indikator_tersedia:
                data_profil[kolom] = pd.to_numeric(
                    data_profil[kolom],
                    errors="coerce",
                )

            profil_hitung = (
                data_profil.groupby("Klaster")[indikator_tersedia]
                .mean()
                .reindex(URUTAN_KLASTER)
                .dropna(how="all")
                .reset_index()
            )

            tampilkan_tabel(profil_hitung)

            indikator_pilihan = st.selectbox(
                "Pilih indikator",
                indikator_tersedia,
            )

            fig = px.bar(
                profil_hitung,
                x="Klaster",
                y=indikator_pilihan,
                color="Klaster",
                color_discrete_map=WARNA_KLASTER,
                category_orders={"Klaster": URUTAN_KLASTER},
                title=f"Rata-rata {indikator_pilihan} per Klaster",
            )
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 15. HALAMAN VISUALISASI PCA
# =========================================================

elif menu == "Visualisasi PCA":
    st.title("Visualisasi PCA")

    st.write(
        "PCA digunakan pada dashboard ini untuk memvisualisasikan "
        "persebaran titik kabupaten/kota dalam dua komponen utama. "
        "PCA bukan dasar penentuan jumlah klaster."
    )

    if data_pca.empty:
        st.warning(
            "Data koordinat PCA tidak ditemukan. Pastikan sheet "
            "Koordinat_PCA tersedia di file Excel."
        )
    else:
        kolom_pca_wilayah = cari_kolom(
            data_pca,
            ["Kabupaten/Kota", "Kabupaten Kota", "Nama Wilayah", "Wilayah"],
        )
        kolom_pc1 = cari_kolom(
            data_pca,
            ["PC1", "PCA1", "Komponen 1", "Principal Component 1"],
        )
        kolom_pc2 = cari_kolom(
            data_pca,
            ["PC2", "PCA2", "Komponen 2", "Principal Component 2"],
        )
        kolom_pca_klaster = cari_kolom(
            data_pca,
            ["Klaster", "Cluster", "Label Klaster"],
        )

        if kolom_pc1 is None or kolom_pc2 is None:
            st.error(
                "Kolom koordinat PCA tidak terbaca. Pastikan sheet PCA "
                "memiliki kolom PC1 dan PC2."
            )
            st.write("Kolom yang terbaca:", list(data_pca.columns))
        else:
            plot_pca = data_pca.copy()

            plot_pca["PC1"] = pd.to_numeric(
                plot_pca[kolom_pc1],
                errors="coerce",
            )
            plot_pca["PC2"] = pd.to_numeric(
                plot_pca[kolom_pc2],
                errors="coerce",
            )

            if kolom_pca_wilayah:
                plot_pca["Kabupaten/Kota"] = (
                    plot_pca[kolom_pca_wilayah].astype(str)
                )
            else:
                st.warning(
                    "Kolom nama wilayah tidak ditemukan pada sheet PCA."
                )
                plot_pca["Kabupaten/Kota"] = (
                    "Wilayah " + (plot_pca.index + 1).astype(str)
                )

            if kolom_pca_klaster:
                plot_pca["Klaster"] = (
                    plot_pca[kolom_pca_klaster].apply(nama_klaster)
                )
            else:
                plot_pca = plot_pca.merge(
                    hasil[["Kabupaten/Kota", "Klaster"]],
                    on="Kabupaten/Kota",
                    how="left",
                )
                plot_pca["Klaster"] = plot_pca["Klaster"].fillna(
                    "Belum cocok"
                )

            plot_pca = plot_pca.dropna(subset=["PC1", "PC2"])

            if plot_pca.empty:
                st.warning("Tidak ada koordinat PCA numerik untuk ditampilkan.")
            else:
                fig = px.scatter(
                    plot_pca,
                    x="PC1",
                    y="PC2",
                    color="Klaster",
                    color_discrete_map=WARNA_KLASTER,
                    category_orders={"Klaster": URUTAN_KLASTER},
                    hover_name="Kabupaten/Kota",
                    hover_data={
                        "PC1": ":.4f",
                        "PC2": ":.4f",
                    },
                    text="Kabupaten/Kota",
                    title="Persebaran Titik Kabupaten/Kota Berdasarkan PCA",
                )

                fig.update_traces(
                    marker=dict(size=12, line=dict(width=1, color="white")),
                    textposition="top center",
                )

                fig.update_layout(
                    xaxis_title="Komponen Utama 1 (PC1)",
                    yaxis_title="Komponen Utama 2 (PC2)",
                    legend_title="Hasil Klaster",
                )

                st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 16. HALAMAN PETA KLASTER
# =========================================================

elif menu == "Peta Klaster":
    st.title("Peta Persebaran Klaster Sumatera Barat")

    st.write(
        "Peta menunjukkan lokasi kabupaten/kota berdasarkan hasil "
        "klasterisasi. Warna hijau menunjukkan Klaster 1, kuning "
        "Klaster 2, dan merah Klaster 3."
    )

    if gpd is None:
        st.error(
            "GeoPandas belum terpasang. Tambahkan geopandas, pyogrio, "
            "dan shapely pada requirements.txt, lalu lakukan deploy ulang."
        )

    elif not Path(FILE_SHAPEFILE).exists():
        st.error(
            f"File '{FILE_SHAPEFILE}' tidak ditemukan. Pastikan file "
            "shapefile ZIP sudah ada di repositori GitHub yang sama "
            "dengan app.py."
        )

    else:
        try:
            @st.cache_data
            def baca_shapefile(nama_zip):
                with tempfile.TemporaryDirectory() as folder:
                    with zipfile.ZipFile(nama_zip, "r") as arsip:
                        arsip.extractall(folder)

                    daftar_shp = list(Path(folder).rglob("*.shp"))

                    if not daftar_shp:
                        raise FileNotFoundError(
                            "File .shp tidak ditemukan di dalam ZIP."
                        )

                    # Memilih shapefile poligon jika terdapat beberapa file.
                    kandidat_shp = [
                        file for file in daftar_shp
                        if "_PL" in file.stem.upper()
                    ]
                    file_shp = (
                        kandidat_shp[0]
                        if kandidat_shp
                        else daftar_shp[0]
                    )

                    peta_baca = gpd.read_file(file_shp)

                    # Mengembalikan geometri dan atribut agar tetap tersedia
                    # setelah folder sementara dihapus.
                    return peta_baca.to_json()

            geojson_teks = baca_shapefile(FILE_SHAPEFILE)
            peta = gpd.read_file(
                json.loads(geojson_teks)
            ) if False else None

            # Membaca GeoJSON dari hasil fungsi yang telah di-cache.
            peta = gpd.GeoDataFrame.from_features(
                json.loads(geojson_teks)["features"],
                crs="EPSG:4326",
            )

            # Cari kolom nama wilayah pada shapefile.
            kolom_peta = cari_kolom(
                peta,
                [
                    "kab_kota",
                    "Kabupaten/Kota",
                    "NAMOBJ",
                    "WADMKK",
                    "NAME_2",
                ],
            )

            if kolom_peta is None:
                st.error(
                    "Kolom nama kabupaten/kota tidak ditemukan di shapefile."
                )
                st.write("Kolom shapefile:", list(peta.columns))

            else:
                peta["Nama_Peta"] = peta[kolom_peta].astype(str).str.strip()
                peta["_kunci"] = peta["Nama_Peta"].apply(kunci_teks)

                hasil_peta = hasil[
                    ["Kabupaten/Kota", "Klaster", "_kunci"]
                ].copy()

                # Menangani variasi penulisan Sawah Lunto/Sawahlunto.
                hasil_peta["_kunci"] = hasil_peta["_kunci"].replace({
                    "kotasawahlunto": "sawahlunto",
                    "sawahlunto": "sawahlunto",
                })
                peta["_kunci"] = peta["_kunci"].replace({
                    "kotasawahlunto": "sawahlunto",
                    "sawahlunto": "sawahlunto",
                })

                data_peta = peta.merge(
                    hasil_peta,
                    on="_kunci",
                    how="left",
                )

                data_peta["Nama Wilayah"] = (
                    data_peta["Kabupaten/Kota"]
                    .fillna(data_peta["Nama_Peta"])
                )
                data_peta["Nama Klaster"] = (
                    data_peta["Klaster"].fillna("Belum cocok")
                )

                data_peta_json = json.loads(data_peta.to_json())

                fig = px.choropleth_map(
                    data_peta,
                    geojson=data_peta_json,
                    locations="_kunci",
                    featureidkey="properties._kunci",
                    color="Nama Klaster",
                    color_discrete_map=WARNA_KLASTER,
                    category_orders={
                        "Nama Klaster": URUTAN_KLASTER,
                    },
                    hover_name="Nama Wilayah",
                    hover_data={
                        "Klaster": True,
                        "_kunci": False,
                        "Nama Klaster": True,
                    },
                    map_style="carto-positron",
                    center={
                        "lat": -0.65,
                        "lon": 100.65,
                    },
                    zoom=5.8,
                    opacity=0.65,
                    title="Peta Persebaran Klaster Sumatera Barat",
                )

                fig.update_traces(
                    marker_line_width=1.2,
                    marker_line_color="#FFFFFF",
                )

                fig.update_layout(
                    margin={"r": 0, "t": 50, "l": 0, "b": 0},
                    legend_title_text="Hasil Klaster",
                    height=650,
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

                wilayah_belum_cocok = data_peta[
                    data_peta["Nama Klaster"] == "Belum cocok"
                ]["Nama_Peta"].tolist()

                if wilayah_belum_cocok:
                    st.warning(
                        "Beberapa nama wilayah pada peta belum cocok "
                        "dengan data hasil klaster: "
                        + ", ".join(wilayah_belum_cocok)
                        + ". Periksa penulisan nama wilayah pada Excel "
                        "dan shapefile."
                    )

                st.subheader("Data Wilayah pada Peta")

                tabel_peta = data_peta[
                    ["Nama Wilayah", "Nama Klaster"]
                ].copy()

                tabel_peta.columns = [
                    "Kabupaten/Kota",
                    "Klaster",
                ]

                tampilkan_tabel(tabel_peta)

        except Exception as error:
            st.error(
                "Peta gagal ditampilkan. Periksa format ZIP shapefile, "
                "nama kolom wilayah, dan dependensi aplikasi."
            )
            st.exception(error)


# =========================================================
# 17. CATATAN KAKI
# =========================================================

st.divider()

st.caption(
    "Dashboard penelitian klasterisasi indikator sosial ekonomi "
    "kabupaten/kota di Provinsi Sumatera Barat menggunakan algoritma "
    "K-Means. Evaluasi jumlah klaster menggunakan Davies-Bouldin Index (DBI)."
)
