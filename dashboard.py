"""
================================================================================
 DASHBOARD PERUSAHAAN KONSTRUKSI - PEMENANG TENDER LPSE (+ OSS + Gapensi)
 v2 - tab-tab di atas, styling lebih menarik, bug deteksi "ada proyek" diperbaiki
================================================================================

INSTALASI (kalau belum ada):
    pip install streamlit pandas openpyxl altair

CARA JALANKAN:
    streamlit run dashboard.py

Taruh file ini di folder yang sama dengan pemenang_gabungan.xlsx (hasil dari
lpse_scraper.py), atau upload filenya langsung lewat tombol di sidebar kalau
lokasinya beda.
"""

import os

import altair as alt
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------------
# KONFIGURASI HALAMAN & STYLING
# ------------------------------------------------------------------------

st.set_page_config(
    page_title="SIDAKON - Sistem Informasi Data Perusahaan Konstruksi",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

WARNA_UTAMA = "#4338CA"
WARNA_UTAMA2 = "#7C3AED"
WARNA_AKSEN = "#F59E0B"
PALET_WARNA = ["#4338CA", "#F59E0B", "#10B981", "#EF4444", "#7C3AED", "#EC4899", "#06B6D4"]

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Poppins', sans-serif;
    }}
    .main {{
        padding-top: 1rem;
    }}
    div[data-testid="stMetric"] {{
        background: linear-gradient(135deg, #ffffff 0%, #f2f0ff 100%);
        border: 1px solid #e4e0fb;
        border-left: 5px solid {WARNA_UTAMA};
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 4px 14px rgba(67, 56, 202, 0.10);
        transition: transform 0.15s ease;
    }}
    div[data-testid="stMetric"]:hover {{
        transform: translateY(-3px);
        box-shadow: 0 8px 20px rgba(67, 56, 202, 0.18);
    }}
    div[data-testid="stMetric"] label {{
        color: #5b5b7a; font-weight: 600; font-size: 0.85rem;
        text-transform: uppercase; letter-spacing: 0.03em;
    }}
    div[data-testid="stMetricValue"] {{
        color: {WARNA_UTAMA}; font-weight: 800;
    }}
    .judul-dashboard {{
        background: linear-gradient(120deg, {WARNA_UTAMA} 0%, {WARNA_UTAMA2} 55%, #EC4899 100%);
        padding: 32px 34px;
        border-radius: 18px;
        color: white;
        margin-bottom: 22px;
        box-shadow: 0 10px 30px rgba(67, 56, 202, 0.25);
        position: relative;
        overflow: hidden;
    }}
    .judul-dashboard h1 {{
        color: white; margin: 0; font-size: 2.3rem; font-weight: 800;
        letter-spacing: 0.02em;
    }}
    .judul-dashboard .subjudul {{
        color: #fbbf24; margin: 2px 0 8px 0; font-weight: 700; font-size: 0.95rem;
        letter-spacing: 0.12em; text-transform: uppercase;
    }}
    .judul-dashboard p {{
        color: #ece9ff; margin: 4px 0 0 0; font-size: 1.02rem;
    }}
    div[data-testid="stTabs"] button[data-baseweb="tab"] {{
        font-weight: 700; font-size: 1.02rem; padding: 10px 4px;
    }}
    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: {WARNA_UTAMA};
        border-bottom-color: {WARNA_UTAMA} !important;
    }}
    [data-testid="stAppViewContainer"] {{
        background:
            radial-gradient(circle at 8% 8%, rgba(124,58,237,0.16) 0%, rgba(124,58,237,0) 30%),
            radial-gradient(circle at 92% 95%, rgba(236,72,153,0.16) 0%, rgba(236,72,153,0) 32%),
            radial-gradient(circle at 92% 5%, rgba(245,158,11,0.10) 0%, rgba(245,158,11,0) 25%),
            linear-gradient(160deg, #f2effe 0%, #e9edfd 35%, #fbeaf3 75%, #fdf3e3 100%);
    }}
    [data-testid="stHeader"] {{
        background: rgba(0,0,0,0);
    }}
    /* PENTING: overflow-y auto di sini WAJIB ada supaya sidebar tetap
       bisa discroll - jangan dihapus/ditimpa kalau menambah style lain */
    section[data-testid="stSidebar"] > div {{
        overflow-y: auto !important;
        max-height: 100vh;
    }}
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #f8f7ff 0%, #f1effc 100%);
        border-right: 1px solid #e4e0fb;
    }}
    section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
        color: {WARNA_UTAMA};
    }}
    div[data-testid="stDataFrame"] {{
        border-radius: 10px; overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    }}
</style>
""", unsafe_allow_html=True)

NAMA_FILE_DEFAULT = ["pemenang_gabungan.xlsx", "pemenang_gabungan.csv",
                     "output/pemenang_gabungan.xlsx", "output/pemenang_gabungan.csv"]


# ------------------------------------------------------------------------
# MEMUAT & MEMBERSIHKAN DATA
# ------------------------------------------------------------------------

@st.cache_data
def muat_data(file_terunggah):
    if file_terunggah is not None:
        if file_terunggah.name.endswith(".csv"):
            df = pd.read_csv(file_terunggah)
        else:
            df = pd.read_excel(file_terunggah)
        return df, file_terunggah.name

    for nama in NAMA_FILE_DEFAULT:
        if os.path.exists(nama):
            if nama.endswith(".csv"):
                df = pd.read_csv(nama)
            else:
                df = pd.read_excel(nama)
            return df, nama

    return None, None


def teks_bersih(series):
    """Ubah kolom jadi teks yang rapi: NaN/None/kosong semua jadi string
    kosong ''. Dipakai supaya baris yang sel-nya kosong di Excel (yang
    kebaca sebagai NaN oleh pandas) tidak salah dianggap 'ada isinya'.
    Pakai fillna() SEBELUM astype(str) - kalau dibalik, NaN di pandas versi
    baru tidak selalu berubah jadi string 'nan' sehingga replace() gagal."""
    return series.fillna("").astype(str).str.strip()


def kolom_angka(df, nama_kolom):
    if nama_kolom not in df.columns:
        return pd.Series([0] * len(df))
    s = df[nama_kolom].astype(str)
    s = s.str.replace(r"[Rp.\s]", "", regex=True).str.replace(",", ".", regex=False)
    s = s.str.extract(r"(\d+\.?\d*)")[0]
    return pd.to_numeric(s, errors="coerce").fillna(0)


# ------------------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------------------

st.sidebar.markdown("## 🏗️ SIDAKON")
st.sidebar.caption("**Si**stem Informasi **Da**ta Perusahaan **Kon**struksi")
st.sidebar.caption("Data pemenang tender LPSE Kota Tanjungpinang, diperkaya data OSS & Gapensi.")

file_upload = st.sidebar.file_uploader(
    "Upload file hasil scraping (opsional)",
    type=["xlsx", "csv"],
    help="Kalau tidak diupload, dashboard otomatis mencari pemenang_gabungan.xlsx/.csv di folder ini.",
)

df, sumber_file = muat_data(file_upload)

if df is None:
    st.error(
        "File data tidak ditemukan. Upload file `pemenang_gabungan.xlsx` atau "
        "`.csv` lewat sidebar, atau taruh file itu di folder yang sama dengan "
        "`dashboard.py` ini."
    )
    st.stop()

# --- Bersihkan kolom penting yang sering dipakai untuk cek "ada isi/tidak" ---
if "kode_paket" in df.columns:
    df["kode_paket"] = teks_bersih(df["kode_paket"])
for kolom in ["nama_pemenang", "nama_paket", "instansi", "alamat", "email", "no_hp"]:
    if kolom in df.columns:
        df[kolom] = teks_bersih(df[kolom])

# kolom bantu: True kalau baris ini benar-benar pemenang tender LPSE
# (bukan sekadar perusahaan tambahan dari OSS/Gapensi yang belum pernah menang)
df["_ada_proyek"] = df["kode_paket"] != "" if "kode_paket" in df.columns else False

st.sidebar.success(f"✅ Data dimuat: **{sumber_file}**")
st.sidebar.caption(f"{len(df):,} baris total · {int(df['_ada_proyek'].sum()):,} baris pemenang tender")

st.sidebar.divider()
st.sidebar.markdown("### 🔍 Filter")

if "instansi" in df.columns:
    daftar_instansi = sorted([i for i in df["instansi"].unique() if i])
    pilih_instansi = st.sidebar.multiselect("Instansi LPSE", daftar_instansi, default=[])
else:
    pilih_instansi = []

if "tahun_anggaran" in df.columns:
    # bersihkan supaya tampil sebagai bilangan bulat ("2026"), bukan "2026.0"
    # - pakai Int64 (nullable integer) karena ada baris kosong (perusahaan
    # tanpa proyek) yang perlu tetap boleh NaN/kosong
    df["tahun_anggaran"] = pd.to_numeric(df["tahun_anggaran"], errors="coerce").astype("Int64")
    daftar_tahun = sorted([int(t) for t in df["tahun_anggaran"].dropna().unique()])
    pilih_tahun = st.sidebar.multiselect("Tahun anggaran", daftar_tahun, default=[])
else:
    pilih_tahun = []

kolom_status_proyek = [c for c in df.columns if c.startswith("status_proyek_")]
filter_status_proyek = {}
for kolom in kolom_status_proyek:
    tahun_label = kolom.replace("status_proyek_", "")
    opsi = sorted(df[kolom].dropna().unique().tolist())
    if opsi:
        pilihan = st.sidebar.selectbox(f"Status proyek {tahun_label}", ["(semua)"] + opsi)
        if pilihan != "(semua)":
            filter_status_proyek[kolom] = pilihan


def urutan_keyakinan(label):
    """Urutkan skor keyakinan dari yang paling meyakinkan ke yang paling
    tidak, bukan urutan abjad."""
    if label.startswith("Sangat yakin"):
        return (0, label)
    if label.startswith("Cukup yakin"):
        return (1, label)
    if label.startswith("Perlu verifikasi"):
        return (2, label)
    return (3, label)


if "skor_keyakinan" in df.columns:
    opsi_keyakinan = sorted(df["skor_keyakinan"].dropna().unique().tolist(), key=urutan_keyakinan)
    pilih_keyakinan = st.sidebar.multiselect("Skor keyakinan data kontak", opsi_keyakinan, default=[])
else:
    pilih_keyakinan = []

kata_kunci = st.sidebar.text_input("🔎 Cari nama perusahaan / proyek")

hanya_pemenang = st.sidebar.checkbox("Hanya tampilkan yang pernah menang tender", value=False)

# ------------------------------------------------------------------------
# TERAPKAN FILTER
# ------------------------------------------------------------------------

df_filtered = df.copy()

if pilih_instansi:
    df_filtered = df_filtered[df_filtered["instansi"].isin(pilih_instansi)]
if pilih_tahun:
    df_filtered = df_filtered[df_filtered["tahun_anggaran"].isin(pilih_tahun)]
for kolom, nilai in filter_status_proyek.items():
    df_filtered = df_filtered[df_filtered[kolom] == nilai]
if pilih_keyakinan:
    df_filtered = df_filtered[df_filtered["skor_keyakinan"].isin(pilih_keyakinan)]
if hanya_pemenang:
    df_filtered = df_filtered[df_filtered["_ada_proyek"]]
if kata_kunci:
    kolom_cari = [c for c in ["nama_pemenang", "nama_paket"] if c in df_filtered.columns]
    if kolom_cari:
        masker = False
        for kolom in kolom_cari:
            masker = masker | df_filtered[kolom].str.contains(kata_kunci, case=False, na=False)
        df_filtered = df_filtered[masker]

df_pemenang = df_filtered[df_filtered["_ada_proyek"]]

# ------------------------------------------------------------------------
# HEADER
# ------------------------------------------------------------------------

st.markdown("""
<div class="judul-dashboard">
    <div class="subjudul">✨ Sistem Informasi Data Perusahaan Konstruksi</div>
    <h1>🏗️ SIDAKON — Tanjungpinang</h1>
    <p>Data pemenang tender LPSE Kota Tanjungpinang, diperkaya dengan data OSS (BPS) dan Gapensi</p>
</div>
""", unsafe_allow_html=True)

tab_ringkasan, tab_grafik, tab_tabel, tab_detail = st.tabs(
    ["📊 Ringkasan", "📈 Grafik", "📋 Data Lengkap", "🔎 Detail Perusahaan"]
)

# ------------------------------------------------------------------------
# TAB 1 - RINGKASAN
# ------------------------------------------------------------------------

with tab_ringkasan:
    kolom1, kolom2, kolom3, kolom4 = st.columns(4)
    kolom1.metric("Total baris (sesuai filter)", f"{len(df_filtered):,}")
    kolom2.metric("Perusahaan unik", f"{df_filtered['nama_pemenang'].nunique():,}"
                  if "nama_pemenang" in df_filtered.columns else "-")
    kolom3.metric("Baris pemenang tender", f"{len(df_pemenang):,}")
    nilai_total = kolom_angka(df_pemenang, "harga_terkoreksi").sum()
    kolom4.metric("Total nilai kontrak", f"Rp {nilai_total:,.0f}")

    st.markdown("")
    kolom_a, kolom_b = st.columns([2, 1])

    with kolom_a:
        st.markdown("##### Sebaran instansi (baris sesuai filter)")
        if "instansi" in df_filtered.columns and not df_filtered.empty:
            data_instansi = df_filtered.groupby("instansi").size().reset_index(name="jumlah")
            data_instansi = data_instansi.sort_values("jumlah", ascending=False)
            chart = alt.Chart(data_instansi).mark_bar(cornerRadiusEnd=4).encode(
                x=alt.X("jumlah:Q", title="Jumlah baris"),
                y=alt.Y("instansi:N", sort="-x", title=""),
                color=alt.value(WARNA_UTAMA),
                tooltip=["instansi", "jumlah"],
            ).properties(height=280)
            st.altair_chart(chart, use_container_width=True)

    with kolom_b:
        if kolom_status_proyek:
            kolom_terbaru = sorted(kolom_status_proyek)[-1]
            tahun_label = kolom_terbaru.replace("status_proyek_", "")
            st.markdown(f"##### Status proyek {tahun_label}")
            data_status = df_filtered[kolom_terbaru].value_counts().reset_index()
            data_status.columns = ["status", "jumlah"]
            chart = alt.Chart(data_status).mark_arc(innerRadius=55).encode(
                theta="jumlah:Q",
                color=alt.Color("status:N", scale=alt.Scale(range=PALET_WARNA), legend=alt.Legend(title="")),
                tooltip=["status", "jumlah"],
            ).properties(height=280)
            st.altair_chart(chart, use_container_width=True)

# ------------------------------------------------------------------------
# TAB 2 - GRAFIK
# ------------------------------------------------------------------------

with tab_grafik:
    kolom_kiri, kolom_kanan = st.columns(2)

    with kolom_kiri:
        st.markdown("##### Jumlah paket per instansi")
        if "instansi" in df_pemenang.columns and not df_pemenang.empty:
            data = df_pemenang.groupby("instansi").size().reset_index(name="jumlah").sort_values("jumlah", ascending=False)
            chart = alt.Chart(data).mark_bar(cornerRadiusEnd=4, color=WARNA_UTAMA).encode(
                x=alt.X("jumlah:Q", title="Jumlah paket"),
                y=alt.Y("instansi:N", sort="-x", title=""),
                tooltip=["instansi", "jumlah"],
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Tidak ada data paket pada filter saat ini.")

    with kolom_kanan:
        st.markdown("##### Jumlah paket per tahun anggaran")
        if "tahun_anggaran" in df_pemenang.columns and not df_pemenang.empty:
            data = df_pemenang.groupby("tahun_anggaran").size().reset_index(name="jumlah")
            data["tahun_anggaran"] = data["tahun_anggaran"].astype(int).astype(str)
            chart = alt.Chart(data).mark_bar(cornerRadiusEnd=4, color=WARNA_AKSEN).encode(
                x=alt.X("tahun_anggaran:N", title="Tahun anggaran"),
                y=alt.Y("jumlah:Q", title="Jumlah paket"),
                tooltip=["tahun_anggaran", "jumlah"],
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Tidak ada data paket pada filter saat ini.")

    if kolom_status_proyek:
        st.markdown("##### Status proyek per tahun")
        kolom_grafik = st.columns(len(kolom_status_proyek))
        for i, kolom in enumerate(kolom_status_proyek):
            with kolom_grafik[i]:
                tahun_label = kolom.replace("status_proyek_", "")
                data_status = df_filtered[kolom].value_counts().reset_index()
                data_status.columns = ["status", "jumlah"]
                if not data_status.empty:
                    chart = alt.Chart(data_status).mark_arc(innerRadius=50).encode(
                        theta="jumlah:Q",
                        color=alt.Color("status:N", scale=alt.Scale(range=PALET_WARNA),
                                        legend=alt.Legend(title=f"Status {tahun_label}")),
                        tooltip=["status", "jumlah"],
                    ).properties(height=280)
                    st.altair_chart(chart, use_container_width=True)

    st.markdown("##### 🏆 Top 10 perusahaan berdasarkan jumlah kemenangan tender")
    if "nama_pemenang" in df_pemenang.columns and not df_pemenang.empty:
        data_top = (
            df_pemenang.groupby("nama_pemenang").size().reset_index(name="jumlah_menang")
            .sort_values("jumlah_menang", ascending=False).head(10)
        )
        chart = alt.Chart(data_top).mark_bar(cornerRadiusEnd=4, color="#10B981").encode(
            x=alt.X("jumlah_menang:Q", title="Jumlah kemenangan"),
            y=alt.Y("nama_pemenang:N", sort="-x", title=""),
            tooltip=["nama_pemenang", "jumlah_menang"],
        ).properties(height=350)
        st.altair_chart(chart, use_container_width=True)
    else:
        st.info("Belum ada data kemenangan tender pada filter saat ini.")

# ------------------------------------------------------------------------
# TAB 3 - DATA LENGKAP
# ------------------------------------------------------------------------

with tab_tabel:
    st.markdown(f"##### 📋 Data lengkap ({len(df_filtered):,} baris)")

    kolom_prioritas = [
        "instansi", "tahun_anggaran", "nama_paket", "nama_pemenang", "alamat",
        "email", "no_hp", "npwp", "harga_penawaran", "harga_terkoreksi",
        "skor_keyakinan", "sumber_data_perusahaan",
    ] + kolom_status_proyek
    kolom_tampil = [k for k in kolom_prioritas if k in df_filtered.columns]
    kolom_sisa = [k for k in df_filtered.columns if k not in kolom_tampil and k != "_ada_proyek"]

    st.dataframe(df_filtered[kolom_tampil + kolom_sisa], use_container_width=True, height=480)

    csv_unduh = df_filtered.drop(columns=["_ada_proyek"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Unduh data yang sedang difilter (CSV)",
        data=csv_unduh,
        file_name="pemenang_gabungan_filtered.csv",
        mime="text/csv",
    )

# ------------------------------------------------------------------------
# TAB 4 - DETAIL PERUSAHAAN
# ------------------------------------------------------------------------

with tab_detail:
    st.markdown("##### 🔎 Lihat detail satu perusahaan")

    def tampil(nilai):
        """Rapikan nilai untuk ditampilkan: NaN/None/kosong jadi '-'."""
        if nilai is None or (isinstance(nilai, float) and pd.isna(nilai)):
            return "-"
        teks = str(nilai).strip()
        if teks == "" or teks.lower() in ("nan", "none", "nat"):
            return "-"
        return teks

    if "nama_pemenang" in df_filtered.columns:
        daftar_perusahaan = sorted([p for p in df_filtered["nama_pemenang"].unique() if p])
        if daftar_perusahaan:
            perusahaan_dipilih = st.selectbox("Pilih perusahaan", daftar_perusahaan)
            detail = df_filtered[df_filtered["nama_pemenang"] == perusahaan_dipilih]
            baris_pertama = detail.iloc[0]

            kolom_a, kolom_b = st.columns(2)
            with kolom_a:
                st.markdown(f"**📍 Alamat:** {tampil(baris_pertama.get('alamat'))}")
                st.markdown(f"**✉️ Email:** {tampil(baris_pertama.get('email'))}")
                st.markdown(f"**📞 No. HP:** {tampil(baris_pertama.get('no_hp'))}")
            with kolom_b:
                st.markdown(f"**🧾 NPWP:** {tampil(baris_pertama.get('npwp'))}")
                st.markdown(f"**✅ Skor keyakinan:** {tampil(baris_pertama.get('skor_keyakinan'))}")
                sumber = tampil(baris_pertama.get("sumber_data_perusahaan"))
                if sumber == "-":
                    # kolom ini cuma terisi untuk perusahaan tambahan dari OSS/Gapensi
                    # yang belum pernah menang tender; kalau kosong berarti datanya
                    # berasal dari LPSE (lihat skor keyakinan untuk rincian sumbernya)
                    sumber = "LPSE"
                st.markdown(f"**🗂️ Sumber data:** {sumber}")

            st.markdown("**Riwayat proyek:**")
            riwayat = detail[detail["_ada_proyek"]].copy()

            # PENTING: situs LPSE memblokir akses langsung ke halaman detail
            # paket kalau bukan dari klik sungguhan di halaman pencarian
            # mereka sendiri (proteksi anti-hotlink). Jadi link langsung ke
            # halaman detail (url_paket) akan selalu kena "Akses Ditolak".
            # Solusinya: arahkan ke halaman pencarian LPSE instansi terkait
            # (yang boleh diakses langsung), dan tampilkan kode paketnya
            # supaya bisa di-paste ke kotak pencarian di sana.
            if "kode_lpse" in riwayat.columns:
                riwayat["cari_di_lpse"] = "https://spse.inaproc.id/" + riwayat["kode_lpse"].astype(str) + "/lelang"

            kolom_riwayat = [c for c in
                              ["tahun_anggaran", "instansi", "nama_paket", "kode_paket",
                               "harga_terkoreksi", "cari_di_lpse"]
                              if c in riwayat.columns]
            if kolom_riwayat and not riwayat.empty:
                st.caption(
                    "⚠️ Situs LPSE tidak mengizinkan link langsung ke halaman detail paket. "
                    "Klik 'Cari di LPSE' untuk membuka portalnya, lalu tempel **Kode Paket** "
                    "di kotak pencarian pada halaman tersebut."
                )
                st.dataframe(
                    riwayat[kolom_riwayat],
                    use_container_width=True,
                    column_config={
                        "cari_di_lpse": st.column_config.LinkColumn("Cari di LPSE", display_text="Buka portal ↗"),
                        "tahun_anggaran": st.column_config.NumberColumn("Tahun", format="%d"),
                        "kode_paket": st.column_config.TextColumn("Kode Paket (copy ini)"),
                    },
                    hide_index=True,
                )
            else:
                st.info("Perusahaan ini belum tercatat memenangkan proyek LPSE pada data yang di-scrape.")
        else:
            st.info("Tidak ada perusahaan pada filter saat ini.")
