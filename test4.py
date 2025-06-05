import streamlit as st
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from io import BytesIO
import zipfile

# Tambah hyperlink aktif ke paragraf
def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
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
    text_elem.text = text
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# Inisialisasi sesi login
if "login_state" not in st.session_state:
    st.session_state.login_state = False
if "username" not in st.session_state:
    st.session_state.username = ""

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
                st.success(f"Selamat datang, {username}!")
            else:
                st.error("Username atau password salah.")

# Halaman utama aplikasi
def show_main_app():
    st.sidebar.markdown(f"👤 Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.experimental_rerun()

    st.title("📄 Generator Surat Massal + Hyperlink Aktif")
    template_file = st.file_uploader("📎 Upload Template Word (.docx)", type="docx")
    data_file = st.file_uploader("📊 Upload Excel Data (.xlsx)", type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.write("📋 Data Ditemukan:")
        st.dataframe(df)

        col_nama = st.selectbox("🧾 Kolom Nama Penyelenggara", df.columns)
        col_link = st.selectbox("🔗 Kolom Short Link", df.columns)
        nama_preview = st.selectbox("🔍 Preview Surat untuk", df[col_nama].unique())

        if nama_preview:
            row = df[df[col_nama] == nama_preview].iloc[0]
            context = {
                "nama_penyelenggara": row[col_nama],
                "short_link": row[col_link]
            }
            doc = DocxTemplate(template_file)
            doc.render(context)
            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)
            st.download_button(f"⬇️ Download Preview ({row[col_nama]})", preview_buf.getvalue(), file_name=f"preview_{row[col_nama]}.docx")

        if st.button("🚀 Generate Semua Surat"):
            output_zip = BytesIO()
            log = []

            with zipfile.ZipFile(output_zip, "w") as zf:
                for _, row in df.iterrows():
                    try:
                        tpl = DocxTemplate(template_file)
                        tpl.render({
                            "nama_penyelenggara": row[col_nama],
                            "short_link": "[short_link]"
                        })
                        temp_buf = BytesIO()
                        tpl.save(temp_buf)
                        temp_buf.seek(0)

                        doc = Document(temp_buf)
                        for p in doc.paragraphs:
                            if "[short_link]" in p.text:
                                parts = p.text.split("[short_link]")
                                p.clear()
                                if parts[0]: p.add_run(parts[0])
                                add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                                if len(parts) > 1: p.add_run(parts[1])

                        for p in doc.paragraphs:
                            for run in p.runs:
                                run.font.name = "Arial"
                                run.font.size = Pt(12)

                        file_stream = BytesIO()
                        doc.save(file_stream)
                        zf.writestr(f"{row[col_nama]}.docx", file_stream.getvalue())
                        log.append({"Nama": row[col_nama], "Status": "✅ Berhasil"})
                    except Exception as e:
                        log.append({"Nama": row[col_nama], "Status": f"❌ Gagal: {str(e)}"})

            st.success("✅ Semua surat berhasil diproses.")
            output_zip.seek(0)
            st.download_button("📦 Download Semua Surat (ZIP)", output_zip.getvalue(), file_name="surat_massal.zip")
            st.dataframe(pd.DataFrame(log))

# Jalankan aplikasi
if st.session_state.login_state:
    show_main_app()
else:
    show_login()
