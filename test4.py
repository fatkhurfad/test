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

# Tambahkan hyperlink aktif ke paragraf
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

# Konversi docx ke teks
def docx_to_text(doc):
    return "\n\n".join([p.text for p in doc.paragraphs])

# Konversi teks ke docx sederhana
def text_to_docx(text):
    doc = Document()
    for para in text.split("\n\n"):
        doc.add_paragraph(para)
    return doc

# Fungsi pratinjau isi docx dalam style sederhana
def render_docx_preview_better(doc):
    st.subheader("📖 Pratinjau Surat Mirip Word")

    html = "<div style='background:#fff; padding:30px; border:1px solid #ddd; border-radius:8px; font-family:Arial, sans-serif; font-size:14px; line-height:1.6; text-align:justify;'>"

    for p in doc.paragraphs:
        if not p.text.strip():
            continue

        runs_html = ""
        for run in p.runs:
            text = run.text.replace("\n", "<br>")
            style = ""
            if run.bold:
                style += "font-weight:bold;"
            if run.italic:
                style += "font-style:italic;"
            runs_html += f"<span style='{style}'>{text}</span>"

        indent = ""
        if p.paragraph_format.first_line_indent:
            indent = f"padding-left: {int(p.paragraph_format.first_line_indent.pt)}pt;"

        html += f"<p style='{indent} margin-bottom:1em;'>{runs_html}</p>"

    html += "</div>"

    st.markdown(html, unsafe_allow_html=True)

# Halaman login
def show_login():
    st.set_page_config(page_title="Generator Surat Hyperlink", layout="centered")
    st.title("📬 Selamat Datang di Aplikasi Surat Massal PMT")
    st.markdown("Silakan login untuk menggunakan aplikasi ini.")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.session_state.username = username
                st.session_state.logout_message = False
                st.rerun()
            else:
                st.error("Username atau password salah.")

# Halaman utama
def show_main_app():
    st.sidebar.success(f"Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.logout_message = True
        st.session_state.login_state = False
        st.session_state.username = ""
        st.rerun()

    st.title("📄 Generator Surat Massal + Hyperlink Aktif")

    template_file = st.file_uploader("📎 Upload Template Word (.docx)", type="docx")
    data_file = st.file_uploader("📊 Upload Excel Data (.xlsx)", type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.write("📋 Pratinjau Data:")
        st.dataframe(df)

        col_nama = st.selectbox("🧾 Kolom Nama", df.columns)
        col_link = st.selectbox("🔗 Kolom Short Link", df.columns)
        nama_preview = st.selectbox("🔍 Pilih Nama untuk Preview", df[col_nama].unique())

        if nama_preview:
            row = df[df[col_nama] == nama_preview].iloc[0]
            tpl = DocxTemplate(template_file)
            tpl.render({"nama_penyelenggara": row[col_nama], "short_link": "[short_link]"})
            temp_buf = BytesIO()
            tpl.save(temp_buf)
            temp_buf.seek(0)

            doc = Document(temp_buf)
            for p in doc.paragraphs:
                if "[short_link]" in p.text:
                    parts = p.text.split("[short_link]")
                    p.clear()
                    if parts[0]:
                        p.add_run(parts[0])
                    add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                    if len(parts) > 1:
                        p.add_run(parts[1])
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for p in doc.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

            # Preview baca isi docx biasa
            doc_text = docx_to_text(doc)

            # Tampilkan editable textarea untuk edit isi surat
            edited_text = st.text_area("✍️ Edit isi surat di sini:", value=doc_text, height=300)

            if st.button("💾 Simpan perubahan dan buat file docx"):
                new_doc = text_to_docx(edited_text)
                buf = BytesIO()
                new_doc.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Surat Hasil Edit", buf.getvalue(), file_name=f"edited_{row[col_nama]}.docx")

        if st.button("🚀 Generate Semua Surat"):
            output_zip = BytesIO()
            log = []

            with zipfile.ZipFile(output_zip, "w") as zf:
                for _, row in df.iterrows():
                    try:
                        tpl = DocxTemplate(template_file)
                        tpl.render({"nama_penyelenggara": row[col_nama], "short_link": "[short_link]"})
                        temp_buf = BytesIO()
                        tpl.save(temp_buf)
                        temp_buf.seek(0)

                        doc = Document(temp_buf)
                        for p in doc.paragraphs:
                            if "[short_link]" in p.text:
                                parts = p.text.split("[short_link]")
                                p.clear()
                                if parts[0]:
                                    p.add_run(parts[0])
                                add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                                if len(parts) > 1:
                                    p.add_run(parts[1])
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        for p in doc.paragraphs:
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

# Routing halaman berdasarkan status login/logout
if "login_state" not in st.session_state:
    st.session_state.login_state = False

if st.session_state.get("logout_message", False):
    st.set_page_config(page_title="Sampai Jumpa!", layout="centered")
    st.title("👋 Terima Kasih!")
    st.markdown("Terima kasih telah menggunakan aplikasi ini.\n\n**See you!**")
    if st.button("🔐 Kembali ke Halaman Login"):
        st.session_state.logout_message = False
        st.rerun()
elif st.session_state.login_state:
    show_main_app()
else:
    show_login()
