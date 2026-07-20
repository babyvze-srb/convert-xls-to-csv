import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
import os

st.set_page_config(
    page_title="Tools Rekap - Convert CSV ke XLSX",
    page_icon="📄",
    layout="wide"
)

# =========================
# FORMAT BULAN INDONESIA
# =========================
bulan_id = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

def format_tanggal_indonesia(tanggal_str):
    dt = datetime.strptime(tanggal_str, "%d-%m-%Y")
    return f"{dt.day} {bulan_id[dt.month]} {dt.year}"


def _is_no_column(col_name: str) -> bool:
    """Deteksi apakah nama kolom merujuk ke kolom 'No' / nomor urut
    (fleksibel: 'No', 'No.', 'NO', 'No Urut', 'Nomor', dll)."""
    normalized = re.sub(r"[^a-z0-9]", "", str(col_name).lower())
    return normalized in ("no", "nourut", "nomor", "nomorurut")


def clean_excess_spaces(col: pd.Series) -> pd.Series:
    """Rapikan spasi berlebihan (mis. 'SJ   0001' -> 'SJ 0001', '  1  ' -> '1')
    pada kolom object, tanpa mengubah nilai kosong/NaN."""
    if col.dtype != "object":
        return col
    return col.apply(
        lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else v
    )


def try_convert_numeric(col: pd.Series) -> pd.Series:
    """
    Kolom bertipe object (string) dicoba dikonversi jadi numerik (kuantitas),
    supaya di Excel tersimpan sebagai angka (int/float), bukan teks.

    Aturan:
    - Mendukung format angka Indonesia (mis. "1.234,50" -> 1234.50)
      maupun format polos (mis. "1234" -> 1234).
    - Kolom hanya dikonversi jika SELURUH nilai (non-kosong) berhasil
      diparse sebagai angka. Kalau ada satu saja nilai yang bukan angka
      (nama, kode berhuruf, alamat, dll), kolom dibiarkan tetap string.
    - Kalau semua nilai berupa bilangan bulat -> dikonversi ke Int64
      (integer nullable, tetap aman kalau ada sel kosong/NaN).
    - Kalau ada nilai desimal -> dikonversi ke float, tetap numerik
      (bukan string), bukan int.
    """
    if col.dtype != "object":
        return col

    original = col.copy()
    stripped = col.astype(str).str.strip()

    # Anggap sel kosong sebagai NaN (tidak menggagalkan deteksi numerik)
    is_blank = stripped.eq("") | stripped.str.lower().eq("nan")

    # Bersihkan format angka Indonesia: "1.234,50" -> "1234.50"
    cleaned = stripped.str.replace(".", "", regex=False)
    cleaned = cleaned.str.replace(",", ".", regex=False)

    numeric_col = pd.to_numeric(cleaned, errors="coerce")

    # Jika ada nilai non-blank yang gagal diparse -> bukan kolom numerik,
    # kembalikan kolom aslinya (string) tanpa perubahan.
    if numeric_col[~is_blank].isna().any():
        return original

    # Semua nilai non-blank berhasil jadi angka
    if (numeric_col.dropna() % 1 == 0).all():
        return numeric_col.astype("Int64")
    return numeric_col


# =====================================================================
# MENU 1: CSV -> XLSX CONVERTER
# =====================================================================
def _process_csv_dataframe(uploaded_file):
    """Baca & bersihkan satu file CSV, kembalikan DataFrame siap diekspor
    (trim spasi, rapikan kolom No, konversi kolom angka, tukar 2 kolom pertama)."""
    df = pd.read_csv(
        uploaded_file,
        sep=",",
        quotechar='"',
        encoding="utf-8",
        engine="python"
    )

    # 1) Trim spasi di semua kolom teks
    df = df.apply(
        lambda col: col.str.strip() if col.dtype == "object" else col
    )

    # 2) Rapikan spasi berlebihan (spasi ganda, dll) khusus kolom "No"
    for c in df.columns:
        if _is_no_column(c):
            df[c] = clean_excess_spaces(df[c])

    # 3) Kolom kuantitas/angka dikonversi ke numerik (int/float),
    #    bukan disimpan sebagai string
    df = df.apply(try_convert_numeric)

    cols = list(df.columns)
    if len(cols) >= 2:
        cols[0], cols[1] = cols[1], cols[0]
        df = df[cols]

    return df


def render_csv_to_xlsx():
    st.title("📄 Convert CSV ke XLSX")
    st.write(
        "Upload file CSV untuk SRB dan/atau SGE, lalu convert jadi satu file "
        "Excel (.xlsx) berisi 2 sheet: **SRB** dan **SGE**."
    )

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_srb = st.file_uploader(
            "Upload File CSV - SRB",
            type=["csv"],
            key="csv_uploader_srb"
        )
    with col_up2:
        uploaded_sge = st.file_uploader(
            "Upload File CSV - SGE",
            type=["csv"],
            key="csv_uploader_sge"
        )

    # File acuan untuk menyusun nama file & tanggal (utamakan SRB, fallback SGE)
    ref_file = uploaded_srb if uploaded_srb is not None else uploaded_sge

    today = datetime.now().strftime("%d-%m-%Y")
    suggested_name = f"template_{today}"

    if ref_file is not None:
        suggested_name = os.path.splitext(ref_file.name)[0]

    st.subheader("Pengaturan Nama File")

    save_mode = st.radio(
        "Pilih metode penyimpanan:",
        ["Custom", "Template"],
        horizontal=True,
        key="save_mode_csv"
    )

    if save_mode == "Custom":
        custom_name = st.text_input(
            "Masukkan nama file",
            value=suggested_name,
            key="custom_name_csv"
        )
        if custom_name.strip():
            file_name = f"{custom_name.strip()}.xlsx"
        else:
            file_name = f"{suggested_name}.xlsx"
    else:
        col1, col2 = st.columns(2)

        with col1:
            jenis_dokumen = st.selectbox(
                "Jenis Dokumen",
                ["PENJUALAN", "INVOICE"],
                key="jenis_dokumen_csv"
            )
        with col2:
            kategori1 = st.selectbox(
                "Produk",
                ["OLI", "LPG"],
                key="kategori1_csv"
            )

        now = datetime.now()
        tanggal_text = f"{now.day} {bulan_id[now.month]} {now.year}"

        if ref_file is not None:
            dates = re.findall(r"\d{2}-\d{2}-\d{4}", ref_file.name)

            if len(dates) >= 2:
                tgl_awal = format_tanggal_indonesia(dates[0])
                tgl_akhir = format_tanggal_indonesia(dates[1])
                if dates[0] == dates[1]:
                    tanggal_text = tgl_awal
                else:
                    tanggal_text = f"{tgl_awal} - {tgl_akhir}"
            elif len(dates) == 1:
                tanggal_text = format_tanggal_indonesia(dates[0])

        # Catatan: "Lokasi" tidak lagi jadi bagian nama file karena satu file
        # xlsx sekarang bisa memuat SRB & SGE sekaligus (dipisah per-sheet).
        file_name = f"{jenis_dokumen} {kategori1} {tanggal_text}.xlsx"

    st.write(f"**Nama file yang akan disimpan:** `{file_name}`")

    if uploaded_srb is not None or uploaded_sge is not None:
        st.info(
            'Jika nama mengandung koma seperti:\n'
            '"PT MAJU, JAYA"\n'
            "pastikan di CSV dibungkus tanda kutip."
        )

        sheets = {}
        errors = {}

        if uploaded_srb is not None:
            try:
                sheets["SRB"] = _process_csv_dataframe(uploaded_srb)
            except Exception as e:
                errors["SRB"] = str(e)

        if uploaded_sge is not None:
            try:
                sheets["SGE"] = _process_csv_dataframe(uploaded_sge)
            except Exception as e:
                errors["SGE"] = str(e)

        for label, msg in errors.items():
            st.error(f"Gagal membaca file CSV {label}: {msg}")

        if sheets:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_name, df in sheets.items():
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
            output.seek(0)

            st.success(f"CSV berhasil dibaca: {', '.join(sheets.keys())}")

            st.download_button(
                label="⬇ Download XLSX",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_csv_xlsx"
            )

            st.subheader("Preview Data")
            preview_tabs = st.tabs(list(sheets.keys()))
            for tab, (sheet_name, df) in zip(preview_tabs, sheets.items()):
                with tab:
                    st.dataframe(df, use_container_width=True)
    else:
        st.info("Upload minimal satu file CSV (SRB dan/atau SGE) untuk mulai.")


# =====================================================================
# JALANKAN APLIKASI
# =====================================================================
render_csv_to_xlsx()
