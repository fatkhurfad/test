import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="📄 Generator Surat Massal", layout="wide")

st.title("📄 Generator Surat Massal - UI/UX Baru")

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

uploaded_template = st.file_uploader("📄 Upload Template Word (.docx)", type="docx")
uploaded_excel = st.file_uploader("📊 Upload Data Excel (.xlsx)", type="xlsx")

if uploaded_template and uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    if len(df.columns) < 2:
        st.warning("❗ Data Excel minimal harus punya 2 kolom.")
        st.stop()

    doc_check = Document(uploaded_template)
    doc_text = "\n".join([p.text for p in doc_check.paragraphs])
    placeholders = list(set(re.findall(r"{{(.*?)}}", doc_text)))

    col1, col2 = st.columns([2, 3])

    with col1:
        st.subheader("🧩 Konfigurasi Placeholder")
        placeholder_mapping = {}
        for ph in placeholders:
            placeholder_mapping[ph] = st.selectbox(f"Pilih kolom untuk {{{{{ph}}}}}", df.columns, key=ph)

        with st.expander("⚙️ Opsi Lanjutan"):
            file_name_format = st.text_input(
                "📝 Format nama file output",
                value="Surat - {{nama_penyelenggara}}"
            )

        mode = st.radio("📌 Pilih Mode", ["🔍 Preview Satu Baris", "📦 Generate Semua"])

    with col2:
        if mode == "🔍 Preview Satu Baris":
            row_index = st.number_input("🔎 Pilih baris untuk preview", min_value=1, max_value=len(df), value=1)
            if st.button("Tampilkan Preview"):
                row = df.iloc[row_index - 1]
                doc = Document(uploaded_template)
                for p in doc.paragraphs:
                    for run in p.runs:
                        for ph, col in placeholder_mapping.items():
                            if f"{{{{{ph}}}}}" in run.text:
                                run.text = run.text.replace(f"{{{{{ph}}}}}", str(row[col]))

                for p in doc.paragraphs:
                    for ph, col in placeholder_mapping.items():
                        if f"{{{{{ph}}}}}" in p.text and "http" in str(row[col]):
                            parts = p.text.split(f"{{{{{ph}}}}}")
                            p.clear()
                            if parts[0]: p.add_run(parts[0])
                            add_hyperlink(p, str(row[col]), str(row[col]))
                            if len(parts) > 1: p.add_run(parts[1])

                for p in doc.paragraphs:
                    for run in p.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(12)

                preview_text = "\n".join([p.text for p in doc.paragraphs])
                st.text_area("📝 Isi Surat Preview:", value=preview_text, height=400)

                preview_buffer = BytesIO()
                doc.save(preview_buffer)
                preview_buffer.seek(0)
                st.download_button(
                    label="📄 Download Preview Surat",
                    data=preview_buffer.getvalue(),
                    file_name=f"preview_{row[placeholder_mapping.get('nama_penyelenggara','preview')]}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        else:
            if st.button("🔄 Generate Semua Surat"):
                output_zip = BytesIO()
                failed = []
                success = 0
                with zipfile.ZipFile(output_zip, "w") as zf:
                    for idx, row in df.iterrows():
                        try:
                            doc = Document(uploaded_template)
                            for p in doc.paragraphs:
                                for run in p.runs:
                                    for ph, col in placeholder_mapping.items():
                                        if f"{{{{{ph}}}}}" in run.text:
                                            run.text = run.text.replace(f"{{{{{ph}}}}}", str(row[col]))

                            for p in doc.paragraphs:
                                for ph, col in placeholder_mapping.items():
                                    if f"{{{{{ph}}}}}" in p.text and "http" in str(row[col]):
                                        parts = p.text.split(f"{{{{{ph}}}}}")
                                        p.clear()
                                        if parts[0]: p.add_run(parts[0])
                                        add_hyperlink(p, str(row[col]), str(row[col]))
                                        if len(parts) > 1: p.add_run(parts[1])

                            for p in doc.paragraphs:
                                for run in p.runs:
                                    run.font.name = "Arial"
                                    run.font.size = Pt(12)

                            filename_raw = file_name_format
                            for ph, col in placeholder_mapping.items():
                                filename_raw = filename_raw.replace(f"{{{{{ph}}}}}", str(row[col]))
                            filename = f"{filename_raw.replace('/', '-')}.docx"

                            buffer = BytesIO()
                            doc.save(buffer)
                            zf.writestr(filename, buffer.getvalue())
                            success += 1
                        except Exception as e:
                            failed.append((idx + 1, str(row[placeholder_mapping.get('nama_penyelenggara', '')]), str(e)))

                st.success(f"✅ {success} surat berhasil dibuat")
                if failed:
                    st.error(f"❌ {len(failed)} surat gagal dibuat")
                    st.dataframe(pd.DataFrame(failed, columns=["Baris", "Nama", "Error"]))

                st.download_button(
                    label="📥 Download ZIP Semua Surat",
                    data=output_zip.getvalue(),
                    file_name="surat_massal_output.zip",
                    mime="application/zip"
                )
