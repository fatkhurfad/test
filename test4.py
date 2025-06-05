import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO
import zipfile
import re
import time

st.set_page_config(page_title='📄 Komdigi Surat Generator', layout='wide')

# 🎨 Branding UI Komdigi
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

st.title("📄 Komdigi Surat Generator")

# 📝 Template Editor Sederhana
default_template = "Yth. {{nama_penyelenggara}},\nSilakan mengunjungi link berikut: {{short_link}}\nTerima kasih."
template_text = st.text_area("📝 Edit Template Surat (gunakan {{nama_penyelenggara}} dan {{short_link}})", value=default_template, height=200)

# 📊 Upload Data
uploaded_excel = st.file_uploader("📊 Upload Data Excel (.xlsx)", type="xlsx")

if uploaded_excel:
    df = pd.read_excel(uploaded_excel)

    if "nama_penyelenggara" in df.columns and "short_link" in df.columns:
        if st.button("🚀 Generate Semua Surat (dari Template)"):
            output_zip = BytesIO()
            success = 0
            start = time.time()

            with zipfile.ZipFile(output_zip, "w") as zf:
                for idx, row in df.iterrows():
                    doc = Document()
                    for line in template_text.split("\n"):
                        replaced = line.replace("{{nama_penyelenggara}}", str(row["nama_penyelenggara"]))
                        replaced = replaced.replace("{{short_link}}", str(row["short_link"]))
                        p = doc.add_paragraph(replaced)
                        run = p.runs[0]
                        run.font.name = "Arial"
                        run.font.size = Pt(12)

                    buffer = BytesIO()
                    doc.save(buffer)
                    filename = f"Surat_{str(row['nama_penyelenggara']).replace('/', '-')}.docx"
                    zf.writestr(filename, buffer.getvalue())
                    success += 1

            st.success(f"✅ {success} surat berhasil dibuat dalam {round(time.time() - start, 2)} detik.")
            st.download_button(
                label="📥 Download ZIP Surat",
                data=output_zip.getvalue(),
                file_name="surat_massal_output.zip",
                mime="application/zip"
            )
    else:
        st.error("❗ Kolom 'nama_penyelenggara' dan 'short_link' wajib ada di Excel.")
