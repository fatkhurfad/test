import streamlit as st
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from io import BytesIO
import zipfile

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

    st.title("📄 Generator Surat Massal + Hyperlink Otomatis")
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
                        context = {
                            "nama_penyelenggara": row[col_nama],
                            "short_link": row[col_link]
                        }
                        tpl = DocxTemplate(template_file)
                        tpl.render(context)
                        file_buf = BytesIO()
                        tpl.save(file_buf)
                        file_buf.seek(0)
                        zf.writestr(f"{row[col_nama]}.docx", file_buf.getvalue())
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
