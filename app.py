
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
import geopandas as gpd


# =========================================================
# 1. KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="Dashboard Klasterisasi Sumatera Barat",
    page_icon="📊",
    layout="wide",
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_SHAPEFILE = "Provinsi Sumatera Barat-KAB_KOTA.zip"
FILE_GAMBAR = "rumah_gadang.png"

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
# 2. TAMPILAN
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
        margin-bottom: 4px;
    }

    .subjudul-dashboard {
        color: #4A6683;
        font-size: 16px;
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

def cari_kolom(data, kandidat):
    """Mencari nama kolom berdasarkan beberapa kemungkinan nama."""
    if data is None or data.empty:
        return None

    kolom_asli = list(data.columns)

    for nama in kandidat:
        for kolom in kolom_asli:
            if str(kolom).strip().lower() == nama.strip().lower():
                return kolom

    for nama in kandidat:
        for kolom in kolom_asli:
            if nama.strip().lower() in str(kolom).strip().lower():
                return kolom

    return None


def kunci_teks(teks):
    """Menyamakan format penulisan nama wilayah."""
    teks = str(teks).strip().lower()
    teks = teks.replace("&", "dan")
    teks = re.sub(r"^(kabupaten|kab\.|kota)\s+", "", teks)
    teks = re.sub(r"[^a-z0-9]", "", teks)
    return teks


def nama_klaster(nilai):
    """Menyeragamkan label klaster."""
    if pd.isna(nilai):
        return "Belum cocok"

    teks = str(nilai).strip()
    angka = re.search(r"\d+", teks)

    if angka:
        nomor = int(angka.group())
        if nomor in [1, 2, 3]:
            return f"Klaster {nomor}"

    return "Belum cocok"


def cari_sheet(semua_sheet, kandidat):
    """Mencari sheet Excel berdasarkan nama."""
    for nama_dicari in kandidat:
        for nama_sheet, data in semua_sheet.items():
            if nama_sheet.strip().lower() == nama_dicari.strip().lower():
                return data.copy()

    for nama_dicari in kandidat:
        for nama_sheet, data in semua_sheet.items():
            nama_normal = nama_sheet.strip().lower().replace(" ", "_")
            if nama_dicari.strip().lower().replace(" ", "_") in nama_normal:
                return data.copy()

    return pd.DataFrame()


def tampilkan_tabel(data, tinggi=420):
    st.dataframe(
        data,
        use_container_width=True,
        height=tinggi,
        hide_index=True,
    )


# =========================================================
# 4. MEMBACA EXCEL
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
        ["Koordinat_PCA", "Koordinat PCA", "Visualisasi_PCA", "PCA"],
    )

    if hasil.empty and semua_sheet:
        hasil = next(iter(semua_sheet.values())).copy()

    return hasil, evaluasi, profil, pca, list(semua_sheet.keys())


if not Path(FILE_EXCEL).exists():
    st.error(
        f"File {FILE_EXCEL} tidak ditemukan. "
        "Pastikan file Excel sudah diunggah ke repositori GitHub "
        "yang sama dengan app.py."
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
    st.error("Data hasil klaster tidak ditemukan dalam file Excel.")
    st.stop()


# =========================================================
# 5. MENYIAPKAN DATA HASIL KLASTER
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
        "Kolom nama wilayah tidak ditemukan. "
        "Periksa nama kolom pada sheet hasil klaster."
    )
    st.write("Kolom yang tersedia:", list(data_asli.columns))
    st.stop()

if kolom_klaster is None:
    st.error(
        "Kolom klaster tidak ditemukan. "
        "Periksa nama kolom Klaster atau Cluster pada Excel."
    )
    st.write("Kolom yang tersedia:", list(data_asli.columns))
    st.stop()

hasil = data_asli.copy()

hasil["Kabupaten/Kota"] = (
    hasil[kolom_wilayah].astype(str).str.strip()
)
hasil["Klaster"] = hasil[kolom_klaster].apply(nama_klaster)
hasil["_kunci"] = hasil["Kabupaten/Kota"].apply(kunci_teks)

hasil = hasil[
    hasil["Kabupaten/Kota"].notna()
    & (hasil["Kabupaten/Kota"] != "")
].copy()

hasil = hasil.drop_duplicates(
    subset="_kunci",
    keep="first",
).reset_index(drop=True)

indikator_tersedia = []

for indikator in INDIKATOR:
    kolom = cari_kolom(hasil, [indikator])
    if kolom is not None and kolom not in indikator_tersedia:
        indikator_tersedia.append(kolom)

jumlah_wilayah = hasil["Kabupaten/Kota"].nunique()
jumlah_indikator = len(indikator_tersedia)


# =========================================================
# 6. MENYIAPKAN EVALUASI DBI
# =========================================================

kolom_k = cari_kolom(
    evaluasi_dbi,
    ["K", "Jumlah Klaster", "Jumlah_Klaster"],
)

kolom_dbi = cari_kolom(
    evaluasi_dbi,
    ["DBI", "Davies Bouldin Index", "Davies-Bouldin Index"],
)

if (
    not evaluasi_dbi.empty
    and kolom_k is not None
    and kolom_dbi is not None
):
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
    ).sort_values("K").reset_index(drop=True)

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
# 7. HEADER DASHBOARD
# =========================================================

kolom_gambar, kolom_judul = st.columns([1, 5])

with kolom_gambar:
    if Path(FILE_GAMBAR).exists():
        st.image(FILE_GAMBAR, use_container_width=True)
    else:
        st.markdown(
            """
            <div style="
                background-color:#DDEBFA;
                border-radius:12px;
                padding:16px;
                text-align:center;
                font-size:45px;">
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
# 8. MENU NAVIGASI
# =========================================================

with st.sidebar:
    st.title("📊 Navigasi")

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
    st.caption("Dashboard penelitian Sumatera Barat")


# =========================================================
# 9. BERANDA
# =========================================================

if menu == "Beranda":

    st.title("Beranda")

    st.write(
        "Dashboard ini menampilkan hasil klasterisasi kabupaten/kota "
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
        )
    else:
        metrik3.metric(
            "Jumlah Klaster",
            hasil["Klaster"].nunique(),
        )

    st.subheader("Tentang Dashboard")

    st.markdown(
        """
        Dashboard menyediakan informasi berikut:

        - **Hasil Klaster:** hasil pengelompokan setiap kabupaten/kota.
        - **Evaluasi DBI:** perbandingan nilai DBI untuk penentuan K.
        - **Distribusi Klaster:** jumlah dan persentase wilayah per klaster.
        - **Profil Klaster:** perbandingan karakteristik indikator.
        - **Visualisasi PCA:** persebaran titik wilayah dalam dua komponen utama.
        - **Peta Klaster:** persebaran geografis hasil pengelompokan.
        """
    )

    st.subheader("Jumlah Wilayah per Klaster")

    ringkasan = (
        hasil.groupby("Klaster")["Kabupaten/Kota"]
        .nunique()
        .reindex(URUTAN_KLASTER, fill_value=0)
        .reset_index()
    )

    ringkasan.columns = ["Klaster", "Jumlah Wilayah"]

    ringkasan = ringkasan[
        ringkasan["Jumlah Wilayah"] > 0
    ].copy()

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
# 10. HASIL KLASTER
# =========================================================

elif menu == "Hasil Klaster":

    st.title("Hasil Klaster")

    st.write(
        "Tabel berikut menampilkan hasil pengelompokan kabupaten/kota "
        "berdasarkan kemiripan indikator sosial ekonomi."
    )

    kolom_filter1, kolom_filter2 = st.columns([2, 1])

    with kolom_filter1:
        pencarian = st.text_input(
            "Cari kabupaten/kota",
            placeholder="Contoh: Agam atau Kota Padang",
        )

    with kolom_filter2:
        daftar_filter = [
            klaster for klaster in URUTAN_KLASTER
            if klaster in hasil["Klaster"].unique()
        ]

        pilihan_klaster = st.selectbox(
            "Filter klaster",
            ["Semua Klaster"] + daftar_filter,
        )

    tampil = hasil.copy()

    if pencarian.strip():
        tampil = tampil[
            tampil["Kabupaten/Kota"].str.contains(
                pencarian.strip(),
                case=False,
                na=False,
            )
        ]

    if pilihan_klaster != "Semua Klaster":
        tampil = tampil[
            tampil["Klaster"] == pilihan_klaster
        ]

    kolom_tampil = ["Kabupaten/Kota", "Klaster"]

    for kolom in indikator_tersedia:
        if kolom not in kolom_tampil:
            kolom_tampil.append(kolom)

    tampil = tampil[kolom_tampil].copy()

    st.write(f"Jumlah wilayah ditampilkan: **{len(tampil)}**")

    tampilkan_tabel(tampil)

    st.download_button(
        "Unduh hasil klaster (CSV)",
        data=tampil.to_csv(index=False).encode("utf-8-sig"),
        file_name="hasil_klaster_sumatera_barat.csv",
        mime="text/csv",
    )


# =========================================================
# 11. EVALUASI DBI
# =========================================================

elif menu == "Evaluasi DBI":

    st.title("Evaluasi Davies-Bouldin Index (DBI)")

    st.write(
        "DBI digunakan untuk mengevaluasi kualitas hasil klasterisasi. "
        "Secara umum, nilai DBI yang lebih rendah menunjukkan hasil "
        "klaster yang lebih baik."
    )

    if evaluasi_dbi.empty:

        st.warning(
            "Data evaluasi DBI belum terbaca. Pastikan sheet Evaluasi_DBI "
            "memiliki kolom K dan DBI."
        )

        st.write("Sheet Excel yang terdeteksi:", nama_sheet)

    else:

        metrik1, metrik2 = st.columns(2)

        metrik1.metric(
            "K Terbaik",
            f"K = {k_optimal}",
        )

        metrik2.metric(
            "DBI Terendah",
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
            f"Nilai DBI terendah pada data evaluasi adalah "
            f"{dbi_terbaik:.4f}, dengan K = {k_optimal}."
        )


# =========================================================
# 12. DISTRIBUSI KLASTER
# =========================================================

elif menu == "Distribusi Klaster":

    st.title("Distribusi Klaster")

    jumlah_per_klaster = (
        hasil.groupby("Klaster")["Kabupaten/Kota"]
        .nunique()
        .reindex(URUTAN_KLASTER, fill_value=0)
        .reset_index()
    )

    jumlah_per_klaster.columns = [
        "Klaster",
        "Jumlah Wilayah",
    ]

    jumlah_per_klaster = jumlah_per_klaster[
        jumlah_per_klaster["Jumlah Wilayah"] > 0
    ].copy()

    jumlah_per_klaster["Persentase (%)"] = (
        jumlah_per_klaster["Jumlah Wilayah"]
        / jumlah_wilayah
        * 100
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
# 13. PROFIL KLASTER
# =========================================================

elif menu == "Profil Klaster":

    st.title("Profil Klaster")

    st.write(
        "Halaman ini digunakan untuk membandingkan karakteristik "
        "indikator sosial ekonomi pada setiap klaster."
    )

    if not profil_klaster.empty:

        st.subheader("Tabel Profil Klaster")
        tampilkan_tabel(profil_klaster)

        kolom_profil_klaster = cari_kolom(
            profil_klaster,
            ["Klaster", "Cluster", "Label Klaster"],
        )

        if kolom_profil_klaster is not None:

            profil_plot = profil_klaster.copy()

            profil_plot["Klaster"] = (
                profil_plot[kolom_profil_klaster].apply(nama_klaster)
            )

            kolom_angka = [
                kolom
                for kolom in profil_plot.columns
                if kolom != kolom_profil_klaster
                and pd.api.types.is_numeric_dtype(profil_plot[kolom])
            ]

            if kolom_angka:

                indikator_pilihan = st.selectbox(
                    "Pilih indikator",
                    kolom_angka,
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
                st.warning(
                    "Tidak ditemukan kolom indikator numerik "
                    "pada sheet profil klaster."
                )

    elif indikator_tersedia:

        st.info(
            "Sheet Profil_Klaster tidak ditemukan. "
            "Profil dihitung dari indikator yang tersedia pada data hasil."
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
            "Pilih indikator untuk grafik",
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

    else:

        st.warning(
            "Data profil dan kolom indikator tidak ditemukan. "
            "Periksa isi file Excel."
        )


# =========================================================
# 14. VISUALISASI PCA
# =========================================================

elif menu == "Visualisasi PCA":

    st.title("Visualisasi PCA")

    st.write(
        "PCA digunakan untuk memvisualisasikan persebaran titik "
        "kabupaten/kota dalam dua komponen utama. PCA pada dashboard "
        "ini hanya untuk visualisasi, bukan untuk menentukan jumlah klaster."
    )

    if data_pca.empty:

        st.warning(
            "Sheet koordinat PCA tidak ditemukan. Pastikan file Excel "
            "memiliki sheet Koordinat_PCA dengan kolom PC1 dan PC2."
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
                "Kolom koordinat PCA tidak ditemukan. "
                "Pastikan terdapat kolom PC1 dan PC2."
            )

            st.write("Kolom yang tersedia:", list(data_pca.columns))

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

            if kolom_pca_wilayah is not None:

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

            if kolom_pca_klaster is not None:

                plot_pca["Klaster"] = (
                    plot_pca[kolom_pca_klaster].apply(nama_klaster)
                )

            elif kolom_pca_wilayah is not None:

                plot_pca = plot_pca.merge(
                    hasil[["Kabupaten/Kota", "Klaster"]],
                    on="Kabupaten/Kota",
                    how="left",
                )

                plot_pca["Klaster"] = (
                    plot_pca["Klaster"].fillna("Belum cocok")
                )

            else:

                plot_pca["Klaster"] = "Belum cocok"

            plot_pca = plot_pca.dropna(
                subset=["PC1", "PC2"]
            )

            if plot_pca.empty:

                st.warning(
                    "Tidak ada koordinat PCA numerik untuk ditampilkan."
                )

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
                    title="Persebaran Titik Wilayah Berdasarkan PCA",
                )

                fig.update_traces(
                    marker=dict(
                        size=12,
                        line=dict(width=1, color="white"),
                    )
                )

                fig.update_layout(
                    xaxis_title="Komponen Utama 1 (PC1)",
                    yaxis_title="Komponen Utama 2 (PC2)",
                    legend_title="Hasil Klaster",
                )

                st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 15. PETA KLASTER
# =========================================================

elif menu == "Peta Klaster":

    st.title("Peta Persebaran Klaster Sumatera Barat")

    st.write(
        "Peta ini menampilkan persebaran geografis kabupaten/kota "
        "berdasarkan hasil klasterisasi."
    )

    st.markdown(
        """
        **Keterangan warna:**
        - 🟢 Hijau: Klaster 1
        - 🟡 Kuning: Klaster 2
        - 🔴 Merah: Klaster 3
        """
    )

    if not Path(FILE_SHAPEFILE).exists():

        st.error(
            f"File {FILE_SHAPEFILE} tidak ditemukan. "
            "Unggah ZIP shapefile ke repositori GitHub."
        )

    else:

        try:

            @st.cache_data
            def baca_geojson_dari_zip(nama_zip):

                with tempfile.TemporaryDirectory() as folder:

                    with zipfile.ZipFile(nama_zip, "r") as arsip:
                        arsip.extractall(folder)

                    daftar_shp = list(Path(folder).rglob("*.shp"))

                    if not daftar_shp:
                        raise FileNotFoundError(
                            "Tidak ditemukan file .shp di dalam ZIP."
                        )

                    kandidat = [
                        file for file in daftar_shp
                        if "_PL" in file.stem.upper()
                    ]

                    file_shp = kandidat[0] if kandidat else daftar_shp[0]

                    peta_baca = gpd.read_file(file_shp)

                    if peta_baca.crs is not None:
                        peta_baca = peta_baca.to_crs("EPSG:4326")

                    return peta_baca.to_json()

            geojson_teks = baca_geojson_dari_zip(FILE_SHAPEFILE)

            geojson_objek = json.loads(geojson_teks)

            peta = gpd.GeoDataFrame.from_features(
                geojson_objek["features"],
                crs="EPSG:4326",
            )

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
                    "Kolom nama kabupaten/kota tidak ditemukan pada shapefile."
                )

                st.write("Kolom shapefile:", list(peta.columns))

            else:

                peta["Nama_Peta"] = (
                    peta[kolom_peta].astype(str).str.strip()
                )

                peta["_kunci"] = peta["Nama_Peta"].apply(kunci_teks)

                hasil_peta = hasil[
                    ["Kabupaten/Kota", "Klaster", "_kunci"]
                ].copy()

                # Menyeragamkan penulisan Kota Sawahlunto.
                hasil_peta["_kunci"] = hasil_peta["_kunci"].replace(
                    {
                        "kotasawahlunto": "sawahlunto",
                        "sawahlunto": "sawahlunto",
                    }
                )

                peta["_kunci"] = peta["_kunci"].replace(
                    {
                        "kotasawahlunto": "sawahlunto",
                        "sawahlunto": "sawahlunto",
                    }
                )

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
                    margin={
                        "r": 0,
                        "t": 50,
                        "l": 0,
                        "b": 0,
                    },
                    legend_title_text="Hasil Klaster",
                    height=650,
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

                belum_cocok = data_peta[
                    data_peta["Nama Klaster"] == "Belum cocok"
                ]["Nama_Peta"].tolist()

                if belum_cocok:
                    st.warning(
                        "Nama wilayah berikut belum cocok dengan data hasil "
                        "klaster: " + ", ".join(belum_cocok)
                        + ". Periksa penulisan nama pada Excel dan shapefile."
                    )

                st.subheader("Daftar Wilayah pada Peta")

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
                "Peta gagal ditampilkan. Periksa file ZIP shapefile, "
                "kolom nama wilayah, dan versi dependensi."
            )

            st.exception(error)


# =========================================================
# 16. CATATAN KAKI
# =========================================================

st.divider()

st.caption(
    "Dashboard klasterisasi indikator sosial ekonomi kabupaten/kota "
    "di Provinsi Sumatera Barat menggunakan algoritma K-Means "
    "dengan evaluasi Davies-Bouldin Index (DBI)."
)
