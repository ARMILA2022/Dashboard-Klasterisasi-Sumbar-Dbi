import json
import re
import tempfile
import unicodedata
import zipfile
from pathlib import Path

import geopandas as gpd
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =========================================================
# KONFIGURASI DASHBOARD
# =========================================================
st.set_page_config(
    page_title="Dashboard Klasterisasi Sumatera Barat",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_EXCEL = "hasil_pengolahan_kmeans_dbi_sumbar.xlsx"
FILE_ZIP_PETA = "Provinsi Sumatera Barat-KAB_KOTA.zip"

WARNA_KLASTER = {
    "1": "#36A269",
    "2": "#F2C94C",
    "3": "#E45B5B",
}
WARNA_TIDAK_DIPILIH = "#DCE4ED"
BIRU_TUA = "#102F50"
BIRU = "#174A7E"
BIRU_AKSI = "#2878B5"

# =========================================================
# TAMPILAN
# =========================================================
st.markdown(
    """
    <style>
    .stApp {
        background: #F4F7FB;
        color: #1E293B;
    }

    [data-testid="stHeader"] {
        background: rgba(244,247,251,.96);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg,#102F50 0%,#174A7E 100%);
    }

    [data-testid="stSidebar"] * {
        color: white;
    }

    [data-testid="stSidebar"] input {
        color: #172033 !important;
    }

    [data-testid="stSidebar"] [data-baseweb="select"] * {
        color: #172033;
    }

    [data-testid="stSidebar"] [data-baseweb="tag"] {
        background: #2878B5;
    }

    .dashboard-header {
        background: linear-gradient(135deg,#102F50 0%,#2878B5 100%);
        color: white;
        padding: 30px;
        border-radius: 18px;
        min-height: 245px;
        box-shadow: 0 8px 24px rgba(16,47,80,.15);
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
        font-size: 30px;
        line-height: 1.25;
        margin: 0 0 15px 0;
        font-weight: 800;
    }

    .dashboard-header p {
        color: #EDF6FF;
        font-size: 15px;
        line-height: 1.7;
        margin: 0;
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
        background: white;
        border: 1px solid #E0EAF4;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 4px 14px rgba(16,47,80,.05);
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
        background: #174A7E;
        color: white;
        border-radius: 8px;
        border: none;
    }

    .stDownloadButton button:hover {
        background: #2878B5;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# FUNGSI BANTU
# =========================================================
def normalisasi_nama(nama):
    """Menyamakan penulisan nama wilayah untuk pencarian dan peta."""
    if pd.isna(nama):
        return ""

    nama = str(nama).strip().lower()
    nama = unicodedata.normalize("NFKD", nama)
    nama = "".join(
        karakter for karakter in nama
        if not unicodedata.combining(karakter)
    )
    nama = re.sub(r"[^a-z0-9\s]", " ", nama)
    nama = re.sub(r"\s+", " ", nama).strip()

    # Kabupaten Agam dan Agam disamakan.
    nama = re.sub(r"^kabupaten\s+", "", nama)
    nama = re.sub(r"^kab\.?\s+", "", nama)

    # Kota tetap dipertahankan agar Kota Solok tidak tertukar
    # dengan Kabupaten Solok.
    nama = nama.replace("kota sawah lunto", "kota sawahlunto")
    nama = nama.replace("sawah lunto", "sawahlunto")

    return nama


def cari_kolom(data, kemungkinan):
    """Mencari nama kolom meskipun penulisannya sedikit berbeda."""
    kolom = {
        str(c).strip().lower(): c
        for c in data.columns
    }

    for nama in kemungkinan:
        if nama.lower() in kolom:
            return kolom[nama.lower()]

    for kolom_asli in data.columns:
        for nama in kemungkinan:
            if nama.lower() in str(kolom_asli).strip().lower():
                return kolom_asli

    return None


def siapkan_wilayah(data):
    data = data.copy()

    pilihan = [
        "Kabupaten/Kota",
        "Kabupaten_Kota",
        "kab_kota",
        "Wilayah",
        "Daerah",
        "Nama Wilayah",
    ]

    kolom = cari_kolom(data, pilihan)

    if kolom is None and (
        data.index.name is not None
        or not isinstance(data.index, pd.RangeIndex)
    ):
        data = data.reset_index()
        kolom = cari_kolom(data, pilihan + ["index"])

    if kolom is not None and kolom != "Kabupaten/Kota":
        data = data.rename(
            columns={kolom: "Kabupaten/Kota"}
        )

    return data


def siapkan_klaster(data):
    data = data.copy()

    kolom = cari_kolom(
        data,
        ["Klaster", "Cluster", "Label Klaster", "Hasil Klaster"],
    )

    if kolom is not None and kolom != "Klaster":
        data = data.rename(columns={kolom: "Klaster"})

    if "Klaster" in data.columns:
        data["Klaster"] = (
            data["Klaster"]
            .astype(str)
            .str.extract(r"(\d+)", expand=False)
        )

    return data


def angka_id(nilai, desimal=4):
    try:
        return (
            f"{float(nilai):,.{desimal}f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except (TypeError, ValueError):
        return "-"


def urut_klaster(values):
    return sorted(
        values,
        key=lambda x: (
            0, int(x)
        ) if str(x).isdigit() else (1, str(x))
    )


# =========================================================
# MEMBACA FILE EXCEL
# =========================================================
if not Path(FILE_EXCEL).exists():
    st.error(
        f"File Excel tidak ditemukan: {FILE_EXCEL}. "
        "Pastikan file Excel berada satu folder dengan app.py."
    )
    st.stop()

try:
    excel = pd.ExcelFile(FILE_EXCEL)
    daftar_sheet = excel.sheet_names

    def baca_sheet(nama, **kwargs):
        if nama in daftar_sheet:
            return pd.read_excel(
                FILE_EXCEL,
                sheet_name=nama,
                **kwargs,
            )
        return pd.DataFrame()

    data_klaster = baca_sheet("Hasil_Klaster")
    data_dbi = baca_sheet("Evaluasi_DBI")
    data_profil = baca_sheet("Profil_Klaster", index_col=0)
    data_pca = baca_sheet("Koordinat_PCA")

except Exception as error:
    st.error(f"Gagal membaca file Excel: {error}")
    st.stop()

if data_klaster.empty:
    st.error(
        "Sheet Hasil_Klaster kosong atau tidak ditemukan. "
        "Periksa nama sheet pada file Excel."
    )
    st.stop()

data_klaster = siapkan_klaster(
    siapkan_wilayah(data_klaster)
)

if (
    "Kabupaten/Kota" not in data_klaster.columns
    or "Klaster" not in data_klaster.columns
):
    st.error(
        "Sheet Hasil_Klaster harus memiliki kolom "
        "Kabupaten/Kota dan Klaster."
    )
    st.stop()

data_klaster = data_klaster.dropna(
    subset=["Kabupaten/Kota", "Klaster"]
).copy()

data_klaster["Kabupaten/Kota"] = (
    data_klaster["Kabupaten/Kota"].astype(str).str.strip()
)

data_klaster["Klaster"] = (
    data_klaster["Klaster"].astype(str).str.strip()
)

data_klaster = data_klaster.drop_duplicates(
    subset=["Kabupaten/Kota"],
    keep="first",
)

if not data_dbi.empty:
    data_dbi.columns = [
        str(kolom).strip()
        for kolom in data_dbi.columns
    ]

if not data_profil.empty:
    data_profil = siapkan_klaster(
        data_profil.reset_index()
    )

if not data_pca.empty:
    data_pca = siapkan_klaster(
        siapkan_wilayah(data_pca)
    )

# =========================================================
# MENENTUKAN K OPTIMAL BERDASARKAN DBI
# =========================================================
kolom_k = cari_kolom(
    data_dbi,
    ["K", "Jumlah Klaster", "Jumlah Cluster"],
)

kolom_dbi = cari_kolom(
    data_dbi,
    ["DBI", "Davies-Bouldin Index", "Davies Bouldin Index"],
)

k_optimal = data_klaster["Klaster"].nunique()
dbi_terbaik = None

if (
    not data_dbi.empty
    and kolom_k is not None
    and kolom_dbi is not None
):
    data_dbi[kolom_k] = pd.to_numeric(
        data_dbi[kolom_k],
        errors="coerce",
    )

    data_dbi[kolom_dbi] = pd.to_numeric(
        data_dbi[kolom_dbi],
        errors="coerce",
    )

    data_dbi = data_dbi.dropna(
        subset=[kolom_k, kolom_dbi]
    )

    if not data_dbi.empty:
        terbaik = data_dbi.loc[
            data_dbi[kolom_dbi].idxmin()
        ]
        k_optimal = int(terbaik[kolom_k])
        dbi_terbaik = float(terbaik[kolom_dbi])

# =========================================================
# HEADER DASHBOARD
col_logo, col_judul = st.columns([1, 5], vertical_alignment="center")

with col_logo:
    st.image(
        "https://commons.wikimedia.org/wiki/Special:FilePath/Rumah_Gadang_Suku_Minang.jpg",
        width=110
    )

with col_judul:
    st.markdown(
        """
        <h1 style="color:#1976D2; margin-bottom:0;">
            🏠 Dashboard Sumbar
        </h1>
        <p style="font-size:18px; margin-top:4px; color:#555;">
            Klasterisasi Indikator Sosial Ekonomi
        </p>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

with kolom_judul:
    st.markdown(
        """
        <div class="dashboard-header">
            <div class="eyebrow">
                SUMATERA BARAT • DATA ANALYTICS
            </div>
            <h1>Dashboard Klasterisasi Sosial Ekonomi</h1>
            <p>
                Eksplorasi karakteristik sosial ekonomi
                kabupaten/kota di Provinsi Sumatera Barat
                menggunakan algoritma K-Means dan evaluasi
                Davies-Bouldin Index (DBI).
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# SIDEBAR: PENCARIAN DAN FILTER
# =========================================================
semua_wilayah = sorted(
    data_klaster["Kabupaten/Kota"].unique()
)

semua_klaster = urut_klaster(
    data_klaster["Klaster"].dropna().unique()
)

with st.sidebar:
    st.markdown("## 🏠 Dashboard Sumbar")
    st.caption("Klasterisasi indikator sosial ekonomi")
    st.markdown("---")

    st.markdown("### 🔎 Cari Kabupaten/Kota")

    kata_cari = st.text_input(
        "Ketik nama wilayah",
        placeholder="Contoh: Agam, Padang, Solok...",
        key="kata_cari",
        help=(
            "Ketik sebagian atau seluruh nama wilayah. "
            "Tabel, grafik, dan peta akan mengikuti pencarian."
        ),
    )

    if kata_cari.strip():
        hasil_pencarian = [
            wilayah
            for wilayah in semua_wilayah
            if kata_cari.strip().casefold()
            in wilayah.casefold()
        ]
    else:
        hasil_pencarian = semua_wilayah

    if kata_cari.strip() and not hasil_pencarian:
        st.warning(
            "Wilayah tidak ditemukan. Periksa ejaan "
            "atau coba kata kunci yang lebih pendek."
        )

    wilayah_dipilih = st.multiselect(
        "Pilih wilayah",
        options=hasil_pencarian,
        default=hasil_pencarian,
        key="pilihan_wilayah",
        help=(
            "Pilih satu atau beberapa wilayah. "
            "Kosongkan pilihan untuk tidak menampilkan wilayah."
        ),
    )

    klaster_dipilih = st.multiselect(
        "Filter klaster",
        options=semua_klaster,
        default=semua_klaster,
        format_func=lambda x: f"Klaster {x}",
        key="pilihan_klaster",
    )

    if st.button("Reset pencarian dan filter", width="stretch"):
        st.session_state["kata_cari"] = ""
        st.session_state["pilihan_wilayah"] = semua_wilayah
        st.session_state["pilihan_klaster"] = semua_klaster
        st.rerun()

    st.markdown("---")
    st.markdown("### Informasi Penelitian")
    st.caption("Metode: K-Means Clustering")
    st.caption("Evaluasi: Davies-Bouldin Index (DBI)")
    st.caption("Wilayah: 19 kabupaten/kota")
    st.caption("Indikator: 9 indikator sosial ekonomi")

# Filter utama untuk semua tab
data_tampil = data_klaster[
    data_klaster["Kabupaten/Kota"].isin(wilayah_dipilih)
    & data_klaster["Klaster"].isin(klaster_dipilih)
].copy()

# =========================================================
# KARTU RINGKASAN
# =========================================================
st.markdown(
    '<div class="section-heading">Ringkasan Hasil Klasterisasi</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-subtitle">'
    'Ringkasan mengikuti pencarian dan filter yang dipilih.'
    '</div>',
    unsafe_allow_html=True,
)

k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "Wilayah Ditampilkan",
    data_tampil["Kabupaten/Kota"].nunique(),
)

k2.metric("Jumlah Indikator", 9)

k3.metric(
    "Klaster Ditampilkan",
    data_tampil["Klaster"].nunique(),
)

k4.metric(
    "K Optimal (DBI)",
    k_optimal,
    delta=(
        f"DBI = {angka_id(dbi_terbaik)}"
        if dbi_terbaik is not None
        else None
    ),
)

# =========================================================
# TAB DASHBOARD
# =========================================================
(
    tab_hasil,
    tab_peta,
    tab_distribusi,
    tab_profil,
    tab_pca,
    tab_dbi,
) = st.tabs(
    [
        "📋 Hasil Klaster",
        "🗺️ Peta Wilayah",
        "📊 Distribusi",
        "📈 Profil Indikator",
        "🔵 Visualisasi PCA",
        "🎯 Evaluasi DBI",
    ]
)

# =========================================================
# TAB 1: HASIL KLASTER
# =========================================================
with tab_hasil:
    st.subheader("Hasil Pengelompokan Kabupaten/Kota")

    if data_tampil.empty:
        st.info(
            "Tidak ada wilayah sesuai pencarian dan filter. "
            "Ubah kata kunci atau pilih kembali filter."
        )
    else:
        tabel = data_tampil.copy()

        st.dataframe(
            tabel,
            width="stretch",
            hide_index=True,
        )

        st.download_button(
            "⬇️ Unduh hasil filter (CSV)",
            data=tabel.to_csv(
                index=False
            ).encode("utf-8-sig"),
            file_name="hasil_klaster_sumatera_barat.csv",
            mime="text/csv",
        )

# =========================================================
# TAB 2: PETA WILAYAH
# =========================================================
with tab_peta:
    st.subheader("Peta Persebaran Klaster Kabupaten/Kota")

    st.caption(
        "Wilayah terpilih ditampilkan berdasarkan klaster. "
        "Arahkan kursor ke peta untuk melihat nama wilayah."
    )

    if not Path(FILE_ZIP_PETA).exists():
        st.error(
            f"File peta tidak ditemukan: {FILE_ZIP_PETA}. "
            "Pastikan ZIP shapefile berada satu folder dengan app.py."
        )

    elif data_tampil.empty:
        st.info(
            "Pilih minimal satu wilayah dan satu klaster "
            "pada bagian filter."
        )

    else:
        try:
            with tempfile.TemporaryDirectory() as folder_sementara:
                with zipfile.ZipFile(FILE_ZIP_PETA) as berkas_zip:
                    berkas_zip.extractall(folder_sementara)

                daftar_shp = list(
                    Path(folder_sementara).rglob("*.shp")
                )

                if not daftar_shp:
                    st.error(
                        "ZIP tidak berisi file SHP. "
                        "Pastikan ZIP mencakup file .shp, .shx, .dbf, dan .prj."
                    )
                else:
                    # Utamakan shapefile poligon kabupaten/kota.
                    shp_path = next(
                        (
                            p for p in daftar_shp
                            if p.name.lower().endswith("_pl.shp")
                        ),
                        daftar_shp[0],
                    )

                    gdf = gpd.read_file(
                        shp_path,
                        engine="pyogrio",
                    )

                    kolom_nama_peta = cari_kolom(
                        gdf,
                        [
                            "kab_kota",
                            "Kabupaten/Kota",
                            "Kabupaten_Kota",
                            "WADMKK",
                            "NAMOBJ",
                        ],
                    )

                    if kolom_nama_peta is None:
                        st.error(
                            "Kolom nama wilayah tidak ditemukan "
                            f"pada shapefile. Kolom yang tersedia: "
                            f"{', '.join(map(str, gdf.columns))}"
                        )
                    else:
                        if gdf.crs is not None:
                            gdf = gdf.to_crs(epsg=4326)

                        gdf["_nama_peta"] = (
                            gdf[kolom_nama_peta]
                            .astype(str)
                            .str.strip()
                        )

                        # Kunci nama dibuat sama untuk data dan peta.
                        # Contoh: Kabupaten Agam dan Agam menjadi "agam".
                        gdf["_kunci"] = gdf["_nama_peta"].apply(
                            normalisasi_nama
                        )

                        dp = data_tampil.copy()
                        dp["_kunci"] = dp[
                            "Kabupaten/Kota"
                        ].apply(normalisasi_nama)

                        # Data klaster digabung ke poligon peta.
                        gdf_join = gdf.merge(
                            dp[
                                [
                                    "_kunci",
                                    "Kabupaten/Kota",
                                    "Klaster",
                                ]
                            ],
                            on="_kunci",
                            how="left",
                            validate="many_to_one",
                        )

                        gdf_join["Klaster"] = (
                            gdf_join["Klaster"]
                            .fillna("Tidak dipilih")
                        )

                        gdf_join = gdf_join.reset_index(drop=True)
                        gdf_join["_map_id"] = (
                            gdf_join.index.astype(str)
                        )

                        # GeoJSON disimpan dalam memori untuk Plotly.
                        geojson_semua = json.loads(
                            gdf_join.to_json()
                        )

                        batas = gdf_join.total_bounds
                        tengah = {
                            "lon": float(
                                (batas[0] + batas[2]) / 2
                            ),
                            "lat": float(
                                (batas[1] + batas[3]) / 2
                            ),
                        }

                        fig_peta = go.Figure()

                        urutan_peta = (
                            urut_klaster(
                                gdf_join.loc[
                                    gdf_join["Klaster"] != "Tidak dipilih",
                                    "Klaster",
                                ].unique()
                            )
                            + ["Tidak dipilih"]
                        )

                        for klaster in urutan_peta:
                            subset = gdf_join[
                                gdf_join["Klaster"] == klaster
                            ]

                            if subset.empty:
                                continue

                            geojson_subset = json.loads(
                                subset.to_json()
                            )

                            warna = WARNA_KLASTER.get(
                                str(klaster),
                                WARNA_TIDAK_DIPILIH,
                            )

                            fig_peta.add_trace(
                                go.Choroplethmap(
                                    geojson=geojson_subset,
                                    locations=subset[
                                        "_map_id"
                                    ].tolist(),
                                    z=[1] * len(subset),
                                    featureidkey="properties._map_id",
                                    colorscale=[
                                        [0, warna],
                                        [1, warna],
                                    ],
                                    showscale=False,
                                    marker_opacity=(
                                        0.85
                                        if klaster != "Tidak dipilih"
                                        else 0.35
                                    ),
                                    marker_line_width=1,
                                    marker_line_color="#FFFFFF",
                                    name=(
                                        f"Klaster {klaster}"
                                        if klaster != "Tidak dipilih"
                                        else "Tidak dipilih"
                                    ),
                                    customdata=subset[
                                        ["_nama_peta", "Klaster"]
                                    ].to_numpy(),
                                    hovertemplate=(
                                        "<b>%{customdata[0]}</b>"
                                        "<br>%{customdata[1]}"
                                        "<extra></extra>"
                                    ),
                                )
                            )

                        fig_peta.update_layout(
                            map={
                                "style": "open-street-map",
                                "center": tengah,
                                "zoom": 6.2,
                            },
                            height=620,
                            margin={
                                "l": 0,
                                "r": 0,
                                "t": 25,
                                "b": 0,
                            },
                            legend={
                                "title": "Hasil Klaster",
                                "orientation": "h",
                                "y": 1.02,
                                "x": 0,
                            },
                            paper_bgcolor="white",
                            plot_bgcolor="white",
                        )

                        st.plotly_chart(
                            fig_peta,
                            width="stretch",
                            config={"scrollZoom": True},
                        )

                        c1, c2, c3 = st.columns(3)
                        c1.markdown("🟢 **Klaster 1**")
                        c2.markdown("🟡 **Klaster 2**")
                        c3.markdown("🔴 **Klaster 3**")

                        tidak_cocok = gdf_join[
                            (gdf_join["Klaster"] == "Tidak dipilih")
                            & (
                                gdf_join["_kunci"].isin(
                                    dp["_kunci"]
                                )
                            )
                        ]["_nama_peta"].tolist()

                        if tidak_cocok:
                            st.warning(
                                "Ada nama wilayah yang belum cocok "
                                "antara peta dan data: "
                                + ", ".join(tidak_cocok)
                            )

        except Exception as error:
            st.error(
                "Peta gagal dimuat. Periksa ZIP shapefile, "
                "nama kolom wilayah, dan dependensi geopandas/pyogrio."
            )
            st.code(str(error))

# =========================================================
# TAB 3: DISTRIBUSI KLASTER
# =========================================================
with tab_distribusi:
    st.subheader("Distribusi Wilayah per Klaster")

    if data_tampil.empty:
        st.info("Tidak ada data untuk ditampilkan.")
    else:
        distribusi = (
            data_tampil.groupby("Klaster")["Kabupaten/Kota"]
            .nunique()
            .reset_index(name="Jumlah Wilayah")
        )

        distribusi["Klaster"] = (
            distribusi["Klaster"].astype(str)
        )
        distribusi["Nama Klaster"] = distribusi[
            "Klaster"
        ].map(lambda x: f"Klaster {x}")

        fig_distribusi = px.bar(
            distribusi,
            x="Nama Klaster",
            y="Jumlah Wilayah",
            color="Klaster",
            text="Jumlah Wilayah",
            color_discrete_map=WARNA_KLASTER,
            category_orders={
                "Klaster": urut_klaster(
                    distribusi["Klaster"].unique()
                )
            },
        )

        fig_distribusi.update_traces(
            textposition="outside",
            cliponaxis=False,
        )

        fig_distribusi.update_layout(
            height=420,
            showlegend=False,
            paper_bgcolor="white",
            plot_bgcolor="white",
            yaxis_title="Jumlah Kabupaten/Kota",
            xaxis_title="",
        )

        st.plotly_chart(
            fig_distribusi,
            width="stretch",
        )

        for klaster in urut_klaster(
            data_tampil["Klaster"].unique()
        ):
            daftar_wilayah = data_tampil.loc[
                data_tampil["Klaster"] == klaster,
                "Kabupaten/Kota",
            ].tolist()

            with st.expander(
                f"Klaster {klaster} — {len(daftar_wilayah)} wilayah"
            ):
                for wilayah in daftar_wilayah:
                    st.write(f"• {wilayah}")

# =========================================================
# TAB 4: PROFIL INDIKATOR
# =========================================================
with tab_profil:
    st.subheader("Profil Indikator Setiap Klaster")

    if data_profil.empty:
        st.warning(
            "Sheet Profil_Klaster tidak tersedia atau kosong."
        )
    else:
        kolom_klaster_profil = cari_kolom(
            data_profil,
            ["Klaster", "Cluster", "Nama Klaster"],
        )

        if kolom_klaster_profil is None:
            st.warning(
                "Kolom Klaster tidak ditemukan pada sheet Profil_Klaster."
            )
        else:
            profil = data_profil.rename(
                columns={kolom_klaster_profil: "Klaster"}
            ).copy()

            profil["Klaster"] = (
                profil["Klaster"]
                .astype(str)
                .str.extract(r"(\d+)", expand=False)
            )

            profil = profil[
                profil["Klaster"].isin(klaster_dipilih)
            ]

            kolom_numerik = [
                kolom
                for kolom in profil.select_dtypes(
                    include="number"
                ).columns
                if kolom != "Klaster"
            ]

            if profil.empty:
                st.info(
                    "Tidak ada profil yang sesuai dengan filter klaster."
                )
            else:
                st.dataframe(
                    profil,
                    width="stretch",
                    hide_index=True,
                )

                if kolom_numerik:
                    indikator = st.selectbox(
                        "Pilih indikator untuk dibandingkan",
                        options=kolom_numerik,
                    )

                    fig_profil = px.bar(
                        profil,
                        x="Klaster",
                        y=indikator,
                        color="Klaster",
                        text_auto=".3f",
                        color_discrete_map=WARNA_KLASTER,
                    )

                    fig_profil.update_layout(
                        height=420,
                        paper_bgcolor="white",
                        plot_bgcolor="white",
                        showlegend=False,
                        xaxis_title="",
                        yaxis_title=indikator,
                    )

                    st.plotly_chart(
                        fig_profil,
                        width="stretch",
                    )

                    st.caption(
                        "Nilai indikator yang lebih tinggi tidak selalu "
                        "menunjukkan kondisi lebih baik. Interpretasi "
                        "bergantung pada jenis indikator."
                    )

# =========================================================
# TAB 5: VISUALISASI PCA
# =========================================================
with tab_pca:
    st.subheader("Visualisasi Persebaran Wilayah dengan PCA")

    st.caption(
        "PCA digunakan untuk memvisualisasikan titik wilayah "
        "dalam dua dimensi. Penentuan klaster tetap dilakukan "
        "menggunakan algoritma K-Means."
    )

    if data_pca.empty:
        st.warning(
            "Sheet Koordinat_PCA belum tersedia atau kosong."
        )
    else:
        kolom_pc1 = cari_kolom(
            data_pca,
            ["PC1", "PCA1", "Principal Component 1", "Komponen 1"],
        )

        kolom_pc2 = cari_kolom(
            data_pca,
            ["PC2", "PCA2", "Principal Component 2", "Komponen 2"],
        )

        kolom_klaster_pca = cari_kolom(
            data_pca,
            ["Klaster", "Cluster", "Label Klaster"],
        )

        kolom_wilayah_pca = cari_kolom(
            data_pca,
            [
                "Kabupaten/Kota",
                "Kabupaten_Kota",
                "kab_kota",
                "Wilayah",
                "Daerah",
            ],
        )

        if (
            kolom_pc1 is None
            or kolom_pc2 is None
            or kolom_klaster_pca is None
        ):
            st.warning(
                "Kolom PC1, PC2, atau Klaster tidak ditemukan "
                "pada sheet Koordinat_PCA."
            )
        else:
            pca = data_pca.copy()

            pca[kolom_pc1] = pd.to_numeric(
                pca[kolom_pc1],
                errors="coerce",
            )
            pca[kolom_pc2] = pd.to_numeric(
                pca[kolom_pc2],
                errors="coerce",
            )

            pca[kolom_klaster_pca] = (
                pca[kolom_klaster_pca]
                .astype(str)
                .str.extract(r"(\d+)", expand=False)
            )

            pca = pca.dropna(
                subset=[
                    kolom_pc1,
                    kolom_pc2,
                    kolom_klaster_pca,
                ]
            )

            pca = pca[
                pca[kolom_klaster_pca].isin(klaster_dipilih)
            ]

            if kolom_wilayah_pca is not None:
                nama_terpilih = {
                    normalisasi_nama(wilayah)
                    for wilayah in wilayah_dipilih
                }

                pca = pca[
                    pca[kolom_wilayah_pca]
                    .apply(normalisasi_nama)
                    .isin(nama_terpilih)
                ]

            if pca.empty:
                st.info(
                    "Tidak ada titik PCA yang sesuai dengan filter."
                )
            else:
                fig_pca = px.scatter(
                    pca,
                    x=kolom_pc1,
                    y=kolom_pc2,
                    color=kolom_klaster_pca,
                    color_discrete_map=WARNA_KLASTER,
                    hover_name=(
                        kolom_wilayah_pca
                        if kolom_wilayah_pca is not None
                        else None
                    ),
                    labels={
                        kolom_pc1: "Komponen Utama 1 (PC1)",
                        kolom_pc2: "Komponen Utama 2 (PC2)",
                        kolom_klaster_pca: "Klaster",
                    },
                )

                fig_pca.update_traces(
                    marker={
                        "size": 13,
                        "line": {
                            "width": 1,
                            "color": "white",
                        },
                    }
                )

                fig_pca.update_layout(
                    height=540,
                    paper_bgcolor="white",
                    plot_bgcolor="white",
                )

                st.plotly_chart(
                    fig_pca,
                    width="stretch",
                )

# =========================================================
# TAB 6: EVALUASI DBI
# =========================================================
with tab_dbi:
    st.subheader("Evaluasi Jumlah Klaster Menggunakan DBI")

    st.caption(
        "Nilai DBI yang lebih kecil secara umum menunjukkan "
        "klaster yang lebih kompak dan terpisah."
    )

    if (
        data_dbi.empty
        or kolom_k is None
        or kolom_dbi is None
    ):
        st.warning(
            "Periksa sheet Evaluasi_DBI. "
            "Kolom K dan DBI harus tersedia."
        )
    else:
        fig_dbi = px.line(
            data_dbi,
            x=kolom_k,
            y=kolom_dbi,
            markers=True,
            labels={
                kolom_k: "Jumlah Klaster (K)",
                kolom_dbi: "Davies-Bouldin Index (DBI)",
            },
        )

        fig_dbi.update_traces(
            line={
                "color": BIRU_AKSI,
                "width": 3,
            },
            marker={
                "size": 10,
                "color": BIRU_TUA,
            },
        )

        fig_dbi.update_layout(
            height=440,
            paper_bgcolor="white",
            plot_bgcolor="white",
        )

        st.plotly_chart(
            fig_dbi,
            width="stretch",
        )

        st.dataframe(
            data_dbi,
            width="stretch",
            hide_index=True,
        )

        if dbi_terbaik is not None:
            st.success(
                f"K terbaik berdasarkan DBI terkecil adalah "
                f"K = {k_optimal}, dengan DBI = "
                f"{angka_id(dbi_terbaik)}."
            )

        st.download_button(
            "⬇️ Unduh evaluasi DBI (CSV)",
            data=data_dbi.to_csv(
                index=False
            ).encode("utf-8-sig"),
            file_name="evaluasi_dbi_sumatera_barat.csv",
            mime="text/csv",
        )

# =========================================================
# FOOTER
# =========================================================
st.markdown(
    """
    <div class="footer-box">
        Dashboard Penelitian · Klasterisasi Indikator Sosial Ekonomi
        Provinsi Sumatera Barat · K-Means & DBI
    </div>
    """,
    unsafe_allow_html=True,
)
