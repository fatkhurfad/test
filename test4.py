import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import zipfile
import re
import time

st.set_page_config(page_title="Generator Surat Otomatis", layout="wide")

# 🎨 Branding Komdigi
st.markdown("""
<style>
.stApp {
    background-color: #f5f8ff;
    font-family: 'Segoe UI', sans-serif;
}
h1, h2, h3 {
    color: #1a237e;
}
.stButton > button {
    background-color: #003366;
    color: white;
    border-radius: 8px;
    padding: 0.6em 1.2em;
    border: none;
    font-weight: bold;
    transition: 0.3s ease;
}
.stButton > button:hover {
    background-color: #001f4d;
}
</style>
""", unsafe_allow_html=True)

# 🔗 Fungsi hyperlink Word
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

# 🔐 Login
if "login_state" not in st.session_state:
    st.session_state.login_state = False

if not st.session_state.login_state:
    st.title("🔐 Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.success("✅ Login berhasil! Silakan lanjutkan ke bawah.")
            else:
                st.error("❌ Username atau password salah.")

# 💼 Aplikasi utama
if st.session_state.login_state:
    st.title("📄 Generator Surat Massal")

    uploaded_template = st.file_uploader("📄 Upload Template Word (.docx)", type="docx")
    uploaded_excel = st.file_uploader("📊 Upload Data Excel (.xlsx)", type="xlsx")

    if uploaded_template and uploaded_excel:
        df = pd.read_excel(uploaded_excel)
        if len(df.columns) < 2:
            st.warning("❗ Excel harus memiliki minimal 2 kolom.")
            st.stop()

        doc_check = Document(uploaded_template)
        placeholders = list(set(re.findall(r"{{(.*?)}}", "\n".join(p.text for p in doc_check.paragraphs)))

        col_nama = st.selectbox("📌 Pilih kolom Nama Penyelenggara", df.columns)
        col_link = st.selectbox("🔗 Pilih kolom untuk Link", df.columns)
        nama_filter = st.selectbox("🔍 Cari Nama Penyelenggara", df[col_nama].unique())

        if nama_filter:
            row = df[df[col_nama] == nama_filter].iloc[0]
            doc = Document(uploaded_template)

            for p in doc.paragraphs:
                for run in p.runs:
                    run.text = run.text.replace("{{nama_penyelenggara}}", str(row[col_nama]))

            for p in doc.paragraphs:
                if "{{short_link}}" in p.text:
                    parts = p.text.split("{{short_link}}")
                    p.clear()
                    if parts[0]: p.add_run(parts[0])
                    add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                    if len(parts) > 1: p.add_run(parts[1])

            for p in doc.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.subheader("📄 Preview Surat")
            st.text_area("Isi Surat", "\n".join(p.text for p in doc.paragraphs), height=300)
            st.download_button("⬇️ Download Preview", buffer.getvalue(), file_name=f"preview_{row[col_nama]}.docx")

        if st.button("🚀 Generate Semua Surat"):
            with st.spinner("Sedang memproses..."):
                output_zip = BytesIO()
                failed, success = [], 0
                start = time.time()

                with zipfile.ZipFile(output_zip, "w") as zf:
                    for idx, row in df.iterrows():
                        try:
                            doc = Document(uploaded_template)
                            for p in doc.paragraphs:
                                for run in p.runs:
                                    run.text = run.text.replace("{{nama_penyelenggara}}", str(row[col_nama]))
                            for p in doc.paragraphs:
                                if "{{short_link}}" in p.text:
                                    parts = p.text.split("{{short_link}}")
                                    p.clear()
                                    if parts[0]: p.add_run(parts[0])
                                    add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                                    if len(parts) > 1: p.add_run(parts[1])
                            for p in doc.paragraphs:
                                for run in p.runs:
                                    run.font.name = "Arial"
                                    run.font.size = Pt(12)
                            fname = f"{str(row[col_nama]).replace('/', '-')}.docx"
                            buf = BytesIO()
                            doc.save(buf)
                            zf.writestr(fname, buf.getvalue())
                            success += 1
                        except Exception as e:
                            failed.append((idx+1, str(row[col_nama]), str(e)))

                st.success(f"✅ {success} surat selesai dalam {round(time.time() - start, 2)} detik.")
                if failed:
                    st.error(f"❌ {len(failed)} surat gagal.")
                    st.dataframe(pd.DataFrame(failed, columns=["Baris", "Nama", "Error"]))
                st.download_button("📥 Download ZIP Surat", output_zip.getvalue(), "surat_massal_output.zip", mime="application/zip")
