import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
import os
import json
import base64

st.set_page_config(
    page_title="Tools Rekap - CSV & Ekstraksi Foto",
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

# =====================================================================
# MENU 1: CSV -> XLSX CONVERTER
# =====================================================================
def render_csv_to_xlsx():
    st.title("📄 Convert CSV ke XLSX")
    st.write("Upload file CSV lalu convert ke Excel (.xlsx)")

    uploaded_file = st.file_uploader(
        "Upload File CSV",
        type=["csv"],
        key="csv_uploader"
    )

    today = datetime.now().strftime("%d-%m-%Y")
    original_filename = ""
    suggested_name = f"template_{today}"

    if uploaded_file is not None:
        original_filename = uploaded_file.name
        suggested_name = os.path.splitext(original_filename)[0]

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
        col1, col2, col3 = st.columns(3)

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
        with col3:
            kategori2 = st.selectbox(
                "Lokasi",
                ["SRB", "SGE"],
                key="kategori2_csv"
            )

        now = datetime.now()
        tanggal_text = f"{now.day} {bulan_id[now.month]} {now.year}"

        if uploaded_file is not None:
            dates = re.findall(r"\d{2}-\d{2}-\d{4}", uploaded_file.name)

            if len(dates) >= 2:
                tgl_awal = format_tanggal_indonesia(dates[0])
                tgl_akhir = format_tanggal_indonesia(dates[1])
                if dates[0] == dates[1]:
                    tanggal_text = tgl_awal
                else:
                    tanggal_text = f"{tgl_awal} - {tgl_akhir}"
            elif len(dates) == 1:
                tanggal_text = format_tanggal_indonesia(dates[0])

        file_name = f"{jenis_dokumen} {kategori1} {kategori2} {tanggal_text}.xlsx"

    st.write(f"**Nama file yang akan disimpan:** `{file_name}`")

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

            df = df.apply(
                lambda col: col.str.strip() if col.dtype == "object" else col
            )

            cols = list(df.columns)
            if len(cols) >= 2:
                cols[0], cols[1] = cols[1], cols[0]
                df = df[cols]

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Data")
            output.seek(0)

            st.success("CSV berhasil dibaca")

            st.download_button(
                label="⬇ Download XLSX",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_csv_xlsx"
            )

            st.subheader("Preview Data")
            st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.error(f"Gagal membaca file CSV: {e}")


# =====================================================================
# MENU 2: EKSTRAKSI TABEL DARI FOTO (AI VISION)
# =====================================================================
COLUMNS = [
    ("tanggal", "Tanggal"),
    ("no_sj", "No. SJ/DO"),
    ("sopir", "Sopir"),
    ("nama_supplier", "Nama Supplier"),
    ("kg50_is", "50KG - IS"),
    ("kg50_ksg", "50KG - KSG"),
    ("kg12_is", "12KG - IS"),
    ("kg12_ksg", "12KG - KSG"),
    ("kg55_is", "5,5KG - IS"),
    ("kg55_ksg", "5,5KG - KSG"),
    ("istgb_50", "IS+TGB 50KG"),
    ("istgb_12", "IS+TGB 12KG"),
    ("istgb_55", "IS+TGB 5,5KG"),
    ("is_total", "Baris Total?"),
]

EXTRACTION_PROMPT = """Kamu melihat foto tabel berjudul "REKAPITULASI PENGELUARAN LPG/BRIGHT GAS".
Tabel punya kolom: NO. SJ/DO, SOPIR, NAMA SUPPLIER, lalu grup kolom "50 KG" (sub-kolom IS, KSG),
grup "12 KG" (sub-kolom IS, KSG), grup "5,5 KG" (sub-kolom IS, KSG), dan kadang grup "IS + TGB"
(sub-kolom 50KG, 12KG, 5,5KG). Ada juga tanggal di judul tabel (contoh: "30 Juni 2026"),
dan baris TOTAL / GRAND TOTAL di bagian bawah.

Baca SEMUA baris pada tabel di foto ini, termasuk baris yang selnya kosong (isi kosong dengan string kosong "").
Jangan lewatkan baris manapun. Perhatikan baik-baik angka yang tertulis, jangan menebak jika tidak yakin -
kalau benar-benar tidak terbaca, isi dengan "?".

Balas HANYA dengan JSON valid (tanpa markdown, tanpa teks lain), dengan struktur persis seperti ini:

{
  "tanggal": "<tanggal yang tertulis di judul tabel, contoh: 30 Juni 2026, atau '' jika tidak ada>",
  "rows": [
    {
      "no_sj": "",
      "sopir": "",
      "nama_supplier": "",
      "kg50_is": "",
      "kg50_ksg": "",
      "kg12_is": "",
      "kg12_ksg": "",
      "kg55_is": "",
      "kg55_ksg": "",
      "istgb_50": "",
      "istgb_12": "",
      "istgb_55": "",
      "is_total": false
    }
  ]
}

Set "is_total": true HANYA untuk baris TOTAL atau GRAND TOTAL (isi nama_supplier dengan "TOTAL" atau "GRAND TOTAL").
Semua nilai angka ditulis sebagai string persis seperti di foto (boleh kosong "").
"""


def call_claude_extract(api_key, image_bytes, media_type, model="claude-sonnet-5"):
    """Panggil Anthropic API dengan gambar, kembalikan dict hasil parse JSON."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    b64data = base64.standard_b64encode(image_bytes).decode("utf-8")

    message = client.messages.create(
        model=model,
        max_tokens=4000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": b64data,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )

    text_parts = [b.text for b in message.content if getattr(b, "type", "") == "text"]
    raw = "\n".join(text_parts).strip()
    raw = re.sub(r"^```(json)?", "", raw.strip())
    raw = re.sub(r"```$", "", raw.strip()).strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"Gagal membaca respons AI sebagai JSON:\n{e}\n\nRespons mentah:\n{raw[:1000]}")


def build_excel_bytes(rows_df):
    """Bangun file excel (BytesIO) dari dataframe hasil ekstraksi/koreksi."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap LPG"

    header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    total_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    display_cols = [c for c in COLUMNS if c[0] != "is_total"]
    headers = [label for _, label in display_cols]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    for _, row in rows_df.iterrows():
        is_total = str(row.get("Baris Total?", "")).strip().upper() == "YA"
        values = [row.get(label, "") for _, label in display_cols]
        ws.append(values)
        r = ws.max_row
        for c in range(1, len(values) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if is_total:
                cell.fill = total_fill
                cell.font = Font(bold=True)

    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions["D"].width = 24
    ws.freeze_panes = "A2"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def render_photo_extraction():
    st.title("📸 Ekstrak Tabel dari Foto (AI)")
    st.write(
        "Upload foto tabel *Rekapitulasi Pengeluaran LPG/Bright Gas* — AI akan "
        "membaca isi tabelnya secara otomatis, hasilnya bisa dikoreksi lalu "
        "diunduh sebagai Excel."
    )

    if "rows_data" not in st.session_state:
        st.session_state.rows_data = []

    with st.expander("⚙ Pengaturan API Key Anthropic", expanded=not bool(st.session_state.get("api_key"))):
        st.markdown(
            "Dapatkan API Key di [console.anthropic.com/settings/keys]"
            "(https://console.anthropic.com/settings/keys). "
            "Key hanya dipakai untuk sesi ini dan tidak disimpan permanen di server."
        )
        api_key_input = st.text_input(
            "Anthropic API Key",
            value=st.session_state.get("api_key", ""),
            type="password",
            key="api_key_field"
        )
        if api_key_input:
            st.session_state.api_key = api_key_input.strip()

    uploaded_photos = st.file_uploader(
        "Upload Foto Tabel (bisa lebih dari satu)",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key="photo_uploader"
    )

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        process_clicked = st.button("🤖 Proses dengan AI", key="process_btn", use_container_width=True)
    with col_b:
        add_blank_clicked = st.button("➕ Tambah Baris Kosong", key="add_blank_btn", use_container_width=True)
    with col_c:
        clear_clicked = st.button("🧹 Bersihkan Semua Hasil", key="clear_btn", use_container_width=True)

    if clear_clicked:
        st.session_state.rows_data = []
        st.rerun()

    if add_blank_clicked:
        blank = {label: "" for _, label in COLUMNS}
        st.session_state.rows_data.append(blank)

    if process_clicked:
        api_key = st.session_state.get("api_key", "").strip()
        if not api_key:
            st.error("Isi API Key Anthropic dulu di bagian pengaturan di atas.")
        elif not uploaded_photos:
            st.warning("Upload minimal satu foto dulu.")
        else:
            progress = st.progress(0, text="Memulai proses...")
            total = len(uploaded_photos)
            for idx, photo in enumerate(uploaded_photos, start=1):
                progress.progress(idx / total, text=f"Memproses foto {idx}/{total}: {photo.name}")
                try:
                    ext = os.path.splitext(photo.name)[1].lower()
                    media_type = {
                        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                        ".png": "image/png", ".webp": "image/webp",
                    }.get(ext, "image/jpeg")

                    image_bytes = photo.read()
                    result = call_claude_extract(api_key, image_bytes, media_type)

                    tanggal = result.get("tanggal", "")
                    for r in result.get("rows", []):
                        row = {label: r.get(key, "") for key, label in COLUMNS if key not in ("tanggal", "is_total")}
                        row["Tanggal"] = tanggal
                        row["Baris Total?"] = "YA" if r.get("is_total") else ""
                        st.session_state.rows_data.append(row)

                except Exception as e:
                    st.error(f"Gagal memproses {photo.name}: {e}")

            progress.progress(1.0, text="Selesai!")
            st.success(f"Selesai memproses {total} foto. Total baris: {len(st.session_state.rows_data)}")

    st.subheader("Hasil Ekstraksi (bisa dikoreksi langsung di tabel)")

    if st.session_state.rows_data:
        display_labels = [label for _, label in COLUMNS]
        df = pd.DataFrame(st.session_state.rows_data)
        for label in display_labels:
            if label not in df.columns:
                df[label] = ""
        df = df[display_labels]

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            key="data_editor_rows"
        )

        st.session_state.rows_data = edited_df.to_dict("records")

        excel_bytes = build_excel_bytes(edited_df)
        st.download_button(
            label="💾 Export ke Excel",
            data=excel_bytes,
            file_name="Rekapitulasi_LPG.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_lpg"
        )
    else:
        st.info("Belum ada data. Upload foto lalu klik 'Proses dengan AI', atau tambah baris kosong secara manual.")


# =====================================================================
# NAVIGASI / MENU UTAMA
# =====================================================================
st.sidebar.title("📚 Menu")
menu = st.sidebar.radio(
    "Pilih fitur:",
    ["📄 Convert CSV ke XLSX", "📸 Ekstrak Tabel dari Foto (AI)"],
    key="main_menu"
)

if menu == "📄 Convert CSV ke XLSX":
    render_csv_to_xlsx()
else:
    render_photo_extraction()
