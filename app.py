import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
import os

st.set_page_config(page_title="CSV ke XLSX Converter", page_icon="📄")

st.title("📄 Convert CSV ke XLSX")
st.write("Upload file CSV lalu convert ke Excel (.xlsx)")

# =========================
# UPLOAD FILE
# =========================
uploaded_file = st.file_uploader(
    "Upload File CSV",
    type=["csv"]
)

today = datetime.now().strftime("%d-%m-%Y")
date_range = today

original_filename = ""
suggested_name = f"template_{today}"

# =========================
# DETECT NAMA + TANGGAL
# =========================
if uploaded_file is not None:
    original_filename = uploaded_file.name

    # Hapus ekstensi .csv
    suggested_name = os.path.splitext(original_filename)[0]

    # Cari tanggal dari nama file
    dates = re.findall(r"\d{2}-\d{2}-\d{4}", original_filename)

    if len(dates) >= 2:
        date_range = f"{dates[0]}_{dates[1]}"
    elif len(dates) == 1:
        date_range = dates[0]

# =========================
# PENGATURAN NAMA FILE
# =========================
st.subheader("Pengaturan Nama File")

save_mode = st.radio(
    "Pilih metode penyimpanan:",
    ["Custom", "Template"],
    horizontal=True
)

if save_mode == "Custom":

    custom_name = st.text_input(
        "Masukkan nama file",
        value=suggested_name
    )

    if custom_name.strip() == "":
        file_name = f"{suggested_name}.xlsx"
    else:
        file_name = f"{custom_name.strip()}.xlsx"

else:
    col1, col2 = st.columns(2)

    with col1:
        kategori1 = st.selectbox(
            "Pilih kategori pertama",
            ["OLI", "LPG"]
        )

    with col2:
        kategori2 = st.selectbox(
            "Pilih kategori kedua",
            ["SRB", "SGE"]
        )

    file_name = f"Penjualan_{kategori1}_{kategori2}_{date_range}.xlsx"

st.write(f"**Nama file yang akan disimpan:** `{file_name}`")

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

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")

        output.seek(0)

        st.success("CSV berhasil dibaca")

        # Download button sinkron dengan input nama
        st.download_button(
            label="⬇ Download XLSX",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Preview Data")
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Gagal membaca file CSV: {e}")
