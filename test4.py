import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import zipfile
from streamlit_cookies_manager import EncryptedCookieManager

# Setup cookie manager
cookies = EncryptedCookieManager(prefix="surat_")
if not cookies.ready():
    st.stop()
cookies.load()

# Restore login from cookie
if cookies.get("username"):
    st.session_state.login_state = True
    st.session_state.username = cookies.get("username")

if "login_state" not in st.session_state:
    st.session_state.login_state = False
if "username" not in st.session_state:
    st.session_state.username = ""

def show_login():
    st.title("🔐 Selamat Datang di Aplikasi Generator Surat")
    st.markdown("Silakan login untuk mulai menggunakan aplikasi.")
    with st.form("login_form"):
        username = st.text_input("👤 Username")
        password = st.text_input("🔒 Password", type="password")
        if st.form_submit_button("Login"):
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.session_state.username = username
                cookies.set("username", username)
                cookies.save()
                st.success("✅ Login berhasil!")
            else:
                st.error("❌ Username atau password salah.")

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
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

def show_main_app():
    st.sidebar.markdown(f"Halo, **{st.session_state.username}** 👋")
    if st.sidebar.button("🔓 Logout"):
        st.session_state.login_state = False
        cookies.delete("username")
        cookies.save()
        st.experimental_rerun()

    nav = st.sidebar.radio("📂 Menu", ["📄 Generator", "📊 Laporan Aktivitas"])

    if nav == "📄 Generator":
        st.title("📄 Generator Surat Massal")
        uploaded_template = st.file_uploader("📄 Upload Template Word (.docx)", type="docx")
        uploaded_excel = st.file_uploader("📊 Upload Data Excel (.xlsx)", type="xlsx")

        if uploaded_template and uploaded_excel:
            df = pd.read_excel(uploaded_excel)
            if len(df.columns) < 2:
                st.warning("❗ File Excel harus memiliki minimal 2 kolom.")
                return

            col_nama = st.selectbox("📌 Kolom Nama Penyelenggara", df.columns)
            col_link = st.selectbox("🔗 Kolom Link", df.columns)
            nama_preview = st.selectbox("🔍 Preview Surat untuk", df[col_nama].unique())

            if nama_preview:
                row = df[df[col_nama] == nama_preview].iloc[0]
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

                buffer_preview = BytesIO()
                doc.save(buffer_preview)
                buffer_preview.seek(0)

                st.subheader("📄 Preview Surat")
                st.text_area("Isi Surat", "\n".join([p.text for p in doc.paragraphs]), height=300)
                st.download_button("⬇️ Download Preview", buffer_preview.getvalue(), file_name=f"preview_{row[col_nama]}.docx")

            if st.button("🚀 Generate Semua Surat"):
                with st.spinner("Membuat semua surat..."):
                    output_zip = BytesIO()
                    failed = []
                    success = 0
                    activity_log = []

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

                                filename = f"{str(row[col_nama]).replace('/', '-')}.docx"
                                buf = BytesIO()
                                doc.save(buf)
                                zf.writestr(filename, buf.getvalue())
                                success += 1
                                activity_log.append({
                                    "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Nama": row[col_nama],
                                    "Status": "Berhasil"
                                })
                            except Exception as e:
                                failed.append((idx + 1, row[col_nama], str(e)))
                                activity_log.append({
                                    "Waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Nama": row[col_nama],
                                    "Status": "Gagal"
                                })

                    st.success(f"✅ {success} surat berhasil dibuat.")
                    if failed:
                        st.error(f"❌ {len(failed)} gagal.")
                        st.dataframe(pd.DataFrame(failed, columns=["Baris", "Nama", "Error"]))

                    st.download_button("📥 Download ZIP", output_zip.getvalue(), "surat_massal_output.zip", mime="application/zip")
                    df_log = pd.DataFrame(activity_log)
                    st.session_state["activity_log"] = df_log

    elif nav == "📊 Laporan Aktivitas":
        st.title("📊 Laporan Aktivitas")
        if "activity_log" in st.session_state:
            st.dataframe(st.session_state["activity_log"])
            log_file = BytesIO()
            st.session_state["activity_log"].to_excel(log_file, index=False)
            log_file.seek(0)
            st.download_button("📥 Download Laporan", log_file.getvalue(), "laporan_aktivitas.xlsx")
        else:
            st.info("Belum ada aktivitas.")

if st.session_state.login_state:
    show_main_app()
else:
    show_login()
