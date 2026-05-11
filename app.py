import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="CSV ke XLSX Converter", page_icon="📄")

st.title("📄 Convert CSV ke XLSX")
st.write("Upload file CSV lalu convert ke Excel (.xlsx)")

# =========================
# PENGATURAN NAMA FILE (ATAS)
# =========================
st.subheader("Pengaturan Nama File")

save_mode = st.radio(
    "Pilih metode penyimpanan:",
    ["Custom", "Template"],
    horizontal=True
)

today = datetime.now().strftime("%d-%m-%Y")

if save_mode == "Custom":
    custom_name = st.text_input("Masukkan nama file")

    if custom_name.strip() == "":
        file_name = f"template_{today}.xlsx"
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

    file_name = f"Penjualan_{kategori1}_{kategori2}_{today}.xlsx"

st.write(f"**Nama file:** `{file_name}`")

uploaded_file = st.file_uploader(
    "Upload File CSV",
    type=["csv"]
)

if uploaded_file is not None:

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

        st.success("CSV berhasil dibaca dan kolom ditukar")

        st.subheader("Preview Data")
        st.dataframe(df)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")

        output.seek(0)

        st.download_button(
            label="⬇ Download XLSX",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Gagal membaca file CSV: {e}")
