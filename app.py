import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
import os

st.set_page_config(
    page_title="CSV ke XLSX Converter",
    page_icon="📄"
)

st.title("📄 Convert CSV ke XLSX")
st.write("Upload file CSV lalu convert ke Excel (.xlsx)")

# =========================
# FORMAT BULAN INDONESIA
# =========================
bulan_id = {
    1: "JANUARI",
    2: "FEBRUARI",
    3: "MARET",
    4: "APRIL",
    5: "MEI",
    6: "JUNI",
    7: "JULI",
    8: "AGUSTUS",
    9: "SEPTEMBER",
    10: "OKTOBER",
    11: "NOVEMBER",
    12: "DESEMBER"
}

def format_tanggal_indonesia(tanggal_str):
    dt = datetime.strptime(tanggal_str, "%d-%m-%Y")
    return f"{dt.day} {bulan_id[dt.month]} {dt.year}"

# =========================
# UPLOAD FILE
# =========================
uploaded_file = st.file_uploader(
    "Upload File CSV",
    type=["csv"]
)

today = datetime.now().strftime("%d-%m-%Y")

original_filename = ""
suggested_name = f"template_{today}"

# =========================
# DETECT NAMA FILE
# =========================
if uploaded_file is not None:
    original_filename = uploaded_file.name
    suggested_name = os.path.splitext(original_filename)[0]

# =========================
# PENGATURAN NAMA FILE
# =========================
st.subheader("Pengaturan Nama File")

save_mode = st.radio(
    "Pilih metode penyimpanan:",
    ["Custom", "Template"],
    horizontal=True
)

# =========================
# CUSTOM
# =========================
if save_mode == "Custom":

    custom_name = st.text_input(
        "Masukkan nama file",
        value=suggested_name
    )

    if custom_name.strip():
        file_name = f"{custom_name.strip()}.xlsx"
    else:
        file_name = f"{suggested_name}.xlsx"

# =========================
# TEMPLATE
# =========================
else:

    col1, col2, col3 = st.columns(3)

    with col1:
        jenis_dokumen = st.selectbox(
            "Jenis Dokumen",
            ["PENJUALAN", "INVOICE"]
        )

    with col2:
        kategori1 = st.selectbox(
            "Produk",
            ["OLI", "LPG"]
        )

    with col3:
        kategori2 = st.selectbox(
            "Lokasi",
            ["SRB", "SGE"]
        )

    # Default tanggal hari ini
    now = datetime.now()
    tanggal_text = (
        f"{now.day} "
        f"{bulan_id[now.month]} "
        f"{now.year}"
    )

    # Ambil tanggal dari nama file CSV
    if uploaded_file is not None:

        dates = re.findall(
            r"\d{2}-\d{2}-\d{4}",
            uploaded_file.name
        )

        if len(dates) >= 2:

            tgl_awal = format_tanggal_indonesia(dates[0])
            tgl_akhir = format_tanggal_indonesia(dates[1])

            # Jika tanggal sama
            if dates[0] == dates[1]:
                tanggal_text = tgl_awal

            # Jika tanggal berbeda
            else:
                tanggal_text = (
                    f"{tgl_awal} - {tgl_akhir}"
                )

        elif len(dates) == 1:

            tanggal_text = format_tanggal_indonesia(
                dates[0]
            )

    file_name = (
        f"{jenis_dokumen} "
        f"{kategori1} "
        f"{kategori2} "
        f"{tanggal_text}.xlsx"
    )

# =========================
# PREVIEW NAMA FILE
# =========================
st.write(
    f"**Nama file yang akan disimpan:** `{file_name}`"
)

# =========================
# PROCESS FILE
# =========================
if uploaded_file is not None:

    st.info(
        'Jika nama mengandung koma seperti:\n'
        '"PT MAJU, JAYA"\n'
        "pastikan di CSV dibungkus tanda kutip."
    )

    try:

        df = pd.read_csv(
            uploaded_file,
            sep=",",
            quotechar='"',
            encoding="utf-8",
            engine="python"
        )

        # Tukar kolom pertama dan kedua
        cols = list(df.columns)

        if len(cols) >= 2:
            cols[0], cols[1] = cols[1], cols[0]
            df = df[cols]

        # Convert ke Excel
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                index=False,
                sheet_name="Data"
            )

        output.seek(0)

        st.success("CSV berhasil dibaca")

        st.download_button(
            label="⬇ Download XLSX",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Preview Data")

        st.dataframe(
            df,
            use_container_width=True
        )

    except Exception as e:
        st.error(
            f"Gagal membaca file CSV: {e}"
        )
