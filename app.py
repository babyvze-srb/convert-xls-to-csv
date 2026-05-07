import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="CSV ke XLSX Converter", page_icon="📄")

st.title("📄 Convert CSV ke XLSX")
st.write("Upload file CSV lalu convert ke Excel (.xlsx)")

uploaded_file = st.file_uploader(
    "Upload File CSV",
    type=["csv"]
)

if uploaded_file is not None:

    st.info(
        "Jika nama mengandung koma seperti:\n"
        '"PT MAJU, JAYA"\n'
        "pastikan di CSV dibungkus tanda kutip."
    )

    try:
        # Baca CSV
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

        # Convert ke Excel
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")

        output.seek(0)

        st.download_button(
            label="⬇ Download XLSX",
            data=output,
            file_name="hasil_convert.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Gagal membaca file CSV: {e}")