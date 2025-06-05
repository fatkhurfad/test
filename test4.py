import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import zipfile

# Tambahkan hyperlink aktif
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
    text_elem.set(qn("xml:space"), "preserve")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# Halaman login
def show_login():
    st.set_page_config(page_title="Generator Surat Hyperlink", layout="centered")
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

    st.title("📄 Generator Surat Massal dari Template + Hyperlink Aktif")

    template_file = st.file_uploader("📎 Upload Template Word (.docx)", type="docx")
    data_file = st.file_uploader("📊 Upload File Excel (.xlsx)", type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.write("📋 Pratinjau Data:")
        st.dataframe(df)

        col_nama = st.selectbox("🧾 Kolom Nama", df.columns)
        col_link = st.selectbox("🔗 Kolom Short Link", df.columns)
        nama_preview = st.selectbox("🔍 Pilih Nama untuk Preview", df[col_nama].unique())

        # Tombol preview
        if nama_preview:
            row = df[df[col_nama] == nama_preview].iloc[0]
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
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            for p in doc.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)

            st.download_button(
                label=f"⬇️ Download Preview Surat ({row[col_nama]})",
                data=preview_buf.getvalue(),
                file_name=f"preview_{row[col_nama]}.docx"
            )

        # Tombol generate semua surat
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
                                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                        for p in doc.paragraphs:
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            for run in p.runs:
                                run.font.name = "Arial"
                                run.font.size = Pt(12)

                        final_buf = BytesIO()
                        doc.save(final_buf)
                        zf.writestr(f"{row[col_nama]}.docx", final_buf.getvalue())
                        log.append({"Nama": row[col_nama], "Status": "✅ Berhasil"})
                    except Exception as e:
                        log.append({"Nama": row[col_nama], "Status": f"❌ Gagal: {str(e)}"})

            st.success("✅ Semua surat berhasil dibuat!")
            output_zip.seek(0)
            st.download_button("📦 Download ZIP Semua Surat", output_zip.getvalue(), file_name="surat_hyperlink.zip")
            st.dataframe(pd.DataFrame(log))

# Inisialisasi sesi
if "login_state" not in st.session_state:
    st.session_state.login_state = False

if st.session_state.login_state:
    show_main_app()
else:
    show_login()
