import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile

# Menambahkan hyperlink aktif ke paragraf
def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rPr.append(sz)

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0000FF")
    rPr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(underline)

    new_run.append(rPr)

    text_elem = OxmlElement("w:t")
    text_elem.set(qn("xml:space"), "preserve")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# Membuat file .docx untuk setiap baris Excel
def generate_documents(df, col_nama, col_link):
    output_zip = BytesIO()
    log = []

    with zipfile.ZipFile(output_zip, "w") as zf:
        for _, row in df.iterrows():
            try:
                doc = Document()
                p = doc.add_paragraph()
                p.add_run(f"Halo {row[col_nama]}, silakan akses tautan berikut: ")
                add_hyperlink(p, "Klik di sini", str(row[col_link]))
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

                buffer = BytesIO()
                doc.save(buffer)
                zf.writestr(f"{row[col_nama]}.docx", buffer.getvalue())
                log.append({"Nama": row[col_nama], "Status": "✅ Berhasil"})
            except Exception as e:
                log.append({"Nama": row[col_nama], "Status": f"❌ Gagal: {str(e)}"})

    return output_zip.getvalue(), log

# Halaman login
def show_login():
    st.set_page_config(page_title="Login | Generator Surat", layout="centered")
    st.title("🔐 Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.session_state.username = username
                st.experimental_rerun()
            else:
                st.error("Username atau password salah.")

# Halaman utama
def show_main_app():
    st.sidebar.success(f"Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.experimental_rerun()

    st.title("📄 Generator Surat Hyperlink")
    data_file = st.file_uploader("📊 Upload Excel (.xlsx)", type="xlsx")

    if data_file:
        df = pd.read_excel(data_file)
        st.write("📋 Pratinjau Data:")
        st.dataframe(df)

        col_nama = st.selectbox("🧾 Kolom Nama", df.columns)
        col_link = st.selectbox("🔗 Kolom Link", df.columns)

        if st.button("🚀 Generate & Download"):
            zip_bytes, log = generate_documents(df, col_nama, col_link)
            st.success("✅ Dokumen berhasil dibuat!")
            st.download_button("📦 Download ZIP", zip_bytes, file_name="surat_hyperlink.zip")
            st.dataframe(pd.DataFrame(log))

# Inisialisasi
if "login_state" not in st.session_state:
    st.session_state.login_state = False

if st.session_state.login_state:
    show_main_app()
else:
    show_login()
