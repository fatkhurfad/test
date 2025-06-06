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
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

# Bahasa (singkat)
LANGUAGES = {
    "id": {
        "welcome": "Selamat Datang di Aplikasi Surat Massal PMT",
        "login": "Silakan login untuk menggunakan aplikasi ini.",
        "username": "Username",
        "password": "Password",
        "login_button": "Login",
        "logout_button": "Logout",
        "dashboard_title": "Dashboard",
        "generate_title": "Generate Surat Massal",
        "choose_language": "Pilih Bahasa / Choose Language",
        "upload_template": "Upload Template Word (.docx) — Drag & Drop atau klik",
        "upload_data": "Upload Data Excel (.xlsx) — Drag & Drop atau klik",
        "select_name_col": "Pilih kolom Nama",
        "select_link_col": "Pilih kolom Link",
        "search_name": "Cari Nama (ketik untuk filter)",
        "select_name_preview": "Pilih Nama untuk Preview",
        "hide_preview": "Sembunyikan Preview",
        "show_preview": "Tampilkan Preview",
        "preview_letter": "📖 Pratinjau Surat",
        "download_preview": "⬇️ Download Preview Surat",
        "generate_all": "Generate Semua Surat",
        "processing_letters": "Sedang memproses surat...",
        "generate_done": "✅ Proses generate selesai!",
        "download_all_zip": "Download Semua Surat (ZIP)",
        "view_log": "Lihat Log Generate",
        "logout_msg": "👋 Terima Kasih!",
        "logout_submsg": "Terima kasih telah menggunakan aplikasi ini.\n\n**See you!**",
        "back_login": "🔐 Kembali ke Halaman Login",
        "login_fail": "Username atau password salah.",
        "upload_first": "Silakan upload template dan data Excel terlebih dahulu.",
    },
    "en": {
        "welcome": "Welcome to PMT Bulk Letter Application",
        "login": "Please login to use this application.",
        "username": "Username",
        "password": "Password",
        "login_button": "Login",
        "logout_button": "Logout",
        "dashboard_title": "Dashboard",
        "generate_title": "Generate Bulk Letters",
        "choose_language": "Select Language / Pilih Bahasa",
        "upload_template": "Upload Word Template (.docx) — Drag & Drop or click",
        "upload_data": "Upload Excel Data (.xlsx) — Drag & Drop or click",
        "select_name_col": "Select Name Column",
        "select_link_col": "Select Link Column",
        "search_name": "Search Name (type to filter)",
        "select_name_preview": "Select Name for Preview",
        "hide_preview": "Hide Preview",
        "show_preview": "Show Preview",
        "preview_letter": "📖 Letter Preview",
        "download_preview": "⬇️ Download Letter Preview",
        "generate_all": "Generate All Letters",
        "processing_letters": "Processing letters...",
        "generate_done": "✅ Generation process complete!",
        "download_all_zip": "Download All Letters (ZIP)",
        "view_log": "View Generation Log",
        "logout_msg": "👋 Thank You!",
        "logout_submsg": "Thank you for using this application.\n\n**See you!**",
        "back_login": "🔐 Back to Login Page",
        "login_fail": "Wrong username or password.",
        "upload_first": "Please upload the template and Excel data first.",
    },
}

def t(key):
    return LANGUAGES.get(st.session_state.lang, LANGUAGES["id"]).get(key, key)

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
                # Render dengan placeholder [short_link]
                tpl.render({"nama_penyelenggara": row[col_name], "short_link": "[short_link]"})
                temp_buf = BytesIO()
                tpl.save(temp_buf)
                temp_buf.seek(0)

                doc = Document(temp_buf)

                # Ganti [short_link] dengan hyperlink aktif (biru & bisa klik)
                for p in doc.paragraphs:
                    if "[short_link]" in p.text:
                        parts = p.text.split("[short_link]")
                        p.clear()
                        if parts[0]:
                            run_before = p.add_run(parts[0])
                            run_before.font.name = "Arial"
                            run_before.font.size = Pt(12)
                        add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                        if len(parts) > 1 and parts[1]:
                            run_after = p.add_run(parts[1])
                            run_after.font.name = "Arial"
                            run_after.font.size = Pt(12)
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
            status_text.text(f"{t('processing_letters')} {idx + 1} / {total}")

    output_zip.seek(0)
    return output_zip, log

def page_generate():
    st.title(t("generate_title"))

    template_file = st.file_uploader(t("upload_template"), type="docx")
    data_file = st.file_uploader(t("upload_data"), type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.success(f"{len(df)} rows loaded successfully")
        st.dataframe(df)

        col_name = st.selectbox(t("select_name_col"), df.columns)
        col_link = st.selectbox(t("select_link_col"), df.columns)

        search_name = st.text_input(t("search_name"), "")
        filtered_names = df[df[col_name].astype(str).str.contains(search_name, case=False, na=False)][col_name].unique()
        selected_name = st.selectbox(t("select_name_preview"), filtered_names)

        if st.session_state.get("show_preview", True) and selected_name:
            row = df[df[col_name] == selected_name].iloc[0]
            tpl = DocxTemplate(template_file)
            # Render tetap pakai placeholder [short_link] untuk preview
            tpl.render({"nama_penyelenggara": row[col_name], "short_link": "[short_link]"})
            temp_buf = BytesIO()
            tpl.save(temp_buf)
            temp_buf.seek(0)

            doc = Document(temp_buf)
            # Preview sederhana, tampilkan teks saja tanpa hyperlink aktif
            preview_text = "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            st.text_area(t("preview_letter"), preview_text, height=300)

            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)

            st.download_button(
                label=f"{t('download_preview')} ({row[col_name]})",
                data=preview_buf.getvalue(),
                file_name=f"preview_{row[col_name]}.docx",
            )

        if st.button(t("generate_all")):
            with st.spinner(t("processing_letters")):
                zip_file, log = generate_letters_with_progress(template_file, df, col_name, col_link)
            st.success(t("generate_done"))
            st.download_button(t("download_all_zip"), zip_file.getvalue(), file_name="surat_massal.zip")
            with st.expander(t("view_log")):
                st.dataframe(pd.DataFrame(log))
    else:
        st.info(t("upload_first"))

def show_login():
    st.title(t("welcome"))
    st.markdown(t("login"))
    with st.form("login_form"):
        username = st.text_input(t("username"))
        password = st.text_input(t("password"), type="password")
        submitted = st.form_submit_button(t("login_button"))
        if submitted:
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.session_state.username = username
                st.session_state.logout_message = False
                st.experimental_rerun()
            else:
                st.error(t("login_fail"))

def show_main_app():
    st.sidebar.success(f"{t('welcome')}, {st.session_state.username}")
    if st.sidebar.button(t("logout_button")):
        st.session_state.logout_message = True
        st.session_state.login_state = False
        st.session_state.username = ""
        st.experimental_rerun()

    st.sidebar.title(t("choose_language"))
    lang = st.sidebar.selectbox(
        "",
        ["id", "en"],
        index=0 if st.session_state.get("lang", "id") == "id" else 1,
        format_func=lambda x: "Indonesia" if x == "id" else "English",
    )
    st.session_state.lang = lang

    st.sidebar.title("Menu")
    page = st.sidebar.radio("Navigasi", [t("dashboard_title"), t("generate_title")])
    if page == t("dashboard_title"):
        st.title(t("dashboard_title"))
        st.write(f"Selamat datang, **{st.session_state.username}**!")
        st.write("Dashboard belum dibuat.")
    else:
        page_generate()

if "login_state" not in st.session_state:
    st.session_state.login_state = False

if "lang" not in st.session_state:
    st.session_state.lang = "id"

if st.session_state.get("logout_message", False):
    st.title(t("logout_msg"))
    st.markdown(t("logout_submsg"))
    if st.button(t("back_login")):
        st.session_state.logout_message = False
        st.experimental_rerun()
elif st.session_state.login_state:
    show_main_app()
else:
    show_login()
