import streamlit as st
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from io import BytesIO
import zipfile

# Inisialisasi session
if "login_state" not in st.session_state:
    st.session_state.login_state = False
if "username" not in st.session_state:
    st.session_state.username = ""

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

def show_main_app():
    st.sidebar.markdown(f"👤 Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.experimental_rerun()

    st.title("📄 Generator Surat Massal (docxtpl)")
    template_file = st.file_uploader("📎 Upload Template Word (.docx) dengan {{placeholder}}", type="docx")
    data_file = st.file_uploader("📊 Upload Data Excel (.xlsx)", type="xlsx")

    if template_file and data_file:
        df = pd.read_excel(data_file)
        st.write("📋 Data:")
        st.dataframe(df)

        col_nama = st.selectbox("📌 Kolom Nama Penyelenggara", df.columns)
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
            st.download_button("⬇️ Download Preview", preview_buf.getvalue(), file_name=f"preview_{row[col_nama]}.docx")

        if st.button("🚀 Generate Semua Surat"):
            output_zip = BytesIO()
            failed = []
            success = 0

            with zipfile.ZipFile(output_zip, "w") as zf:
                for idx, row in df.iterrows():
                    try:
                        doc = Document(uploaded_template)

                        for p in doc.paragraphs:
                            for run in p.runs:
                                if "{{nama_penyelenggara}}" in run.text:
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

                        custom_filename = file_name_format.replace("{{nama_penyelenggara}}", str(row[col_nama]))
                        filename = f"{custom_filename.replace('/', '-')}.docx"

                        buffer = BytesIO()
                        doc.save(buffer)
                        zf.writestr(filename, buffer.getvalue())
                        success += 1
                    except Exception as e:
                        failed.append((idx + 1, str(row[col_nama]), str(e)))
                st.success("✅ Surat selesai dibuat.")
                st.download_button("📦 Download Semua Surat (ZIP)", zip_buffer.getvalue(), file_name="surat_massal.zip")
                st.dataframe(pd.DataFrame(log))

if st.session_state.login_state:
    show_main_app()
else:
    show_login()
