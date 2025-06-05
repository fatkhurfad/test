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

# Fungsi tambah hyperlink aktif di dokumen Word
def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
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

# Fungsi render preview surat seperti Word (style dasar)
def render_docx_preview_better(doc):
    st.subheader("📖 Pratinjau Surat Mirip Word")

    html = (
        "<div style='background:#fff; padding:30px; border:1px solid #ddd; "
        "border-radius:8px; font-family:Arial, sans-serif; font-size:14px; "
        "line-height:1.6; text-align:justify;'>"
    )

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

# Fungsi generate surat batch dengan progress bar
def generate_letters_with_progress(template_file, df, col_name, col_link):
    output_zip = BytesIO()
    log = []

    with zipfile.ZipFile(output_zip, "w") as zf:
        total = len(df)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, row in df.iterrows():
            try:
                tpl = DocxTemplate(template_file)
                tpl.render(
                    {"nama_penyelenggara": row[col_name], "short_link": "[short_link]"}
                )
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
                zf.writestr(f"{row[col_name]}.docx", final_buf.getvalue())
                log.append({"Nama": row[col_name], "Status": "✅ Berhasil"})
            except Exception as e:
                log.append({"Nama": row[col_name], "Status": f"❌ Gagal: {str(e)}"})

            progress = int((idx + 1) / total * 100)
            progress_bar.progress(progress)
            status_text.text(f"Memproses surat ke-{idx + 1} dari {total}...")

    output_zip.seek(0)
    return output_zip, log

# Halaman Dashboard
def page_home():
    st.title("🏠 Dashboard")
    st.write("Selamat datang di aplikasi Surat Massal PMT versi canggih!")

# Halaman Generate Surat
def page_generate():
    st.title("🚀 Generate Surat Massal")

    template_file = st.file_uploader("Upload Template Word (.docx)", type="docx")
    data_file = st.file_uploader("Upload Data Excel (.xlsx)", type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.dataframe(df)

        col_name = st.selectbox("Pilih kolom Nama", df.columns)
        col_link = st.selectbox("Pilih kolom Link", df.columns)

        # Preview surat per penerima
        nama_preview = st.selectbox("Pilih Nama untuk Preview", df[col_name].unique())
        if nama_preview:
            row = df[df[col_name] == nama_preview].iloc[0]
            tpl = DocxTemplate(template_file)
            tpl.render({"nama_penyelenggara": row[col_name], "short_link": "[short_link]"})
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

            render_docx_preview_better(doc)

            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)

            st.download_button(
                label=f"⬇️ Download Preview Surat ({row[col_name]})",
                data=preview_buf.getvalue(),
                file_name=f"preview_{row[col_name]}.docx",
            )

        if st.button("Generate Semua Surat"):
            zip_file, log = generate_letters_with_progress(
                template_file, df, col_name, col_link
            )
            st.success("✅ Proses generate selesai!")
            st.download_button(
                "Download Semua Surat (ZIP)", zip_file.getvalue(), file_name="surat_massal.zip"
            )
            st.dataframe(pd.DataFrame(log))

# Halaman Login
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

# Routing utama dan halaman logout
def show_main_app():
    st.sidebar.success(f"Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.logout_message = True
        st.session_state.login_state = False
        st.session_state.username = ""
        st.rerun()

    st.sidebar.title("Menu")
    page = st.sidebar.radio("Navigasi", ["Dashboard", "Generate Surat"])

    if page == "Dashboard":
        page_home()
    elif page == "Generate Surat":
        page_generate()

# Entry point
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
