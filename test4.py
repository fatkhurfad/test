import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import zipfile

st.set_page_config(page_title="Generator Surat Dinamis", layout="centered")
st.title("📄 Generator Surat Massal Dinamis + Preview")

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

uploaded_template = st.file_uploader("Upload Template Word (.docx)", type="docx")
uploaded_excel = st.file_uploader("Upload Data Excel (.xlsx)", type="xlsx")

if uploaded_template and uploaded_excel:
    df = pd.read_excel(uploaded_excel)

    if len(df.columns) >= 2:
        st.subheader("📌 Pilih kolom untuk mengganti placeholder:")
        col_nama = st.selectbox("Ganti {{nama}} dengan:", df.columns)
        col_link = st.selectbox("Ganti {{link}} dengan:", df.columns)

        if st.button("🔍 Preview Surat Pertama"):
            row = df.iloc[0]
            doc = Document(uploaded_template)

            for p in doc.paragraphs:
                for run in p.runs:
                    if "{{nama}}" in run.text:
                        run.text = run.text.replace("{{nama}}", str(row[col_nama]))

            for p in doc.paragraphs:
                if "{{link}}" in p.text:
                    parts = p.text.split("{{link}}")
                    p.clear()
                    if parts[0]: p.add_run(parts[0])
                    add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                    if len(parts) > 1: p.add_run(parts[1])

            for p in doc.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

            preview_buffer = BytesIO()
            doc.save(preview_buffer)
            preview_buffer.seek(0)
            st.download_button(
                label="⬇️ Download Preview Surat (1 Data)",
                data=preview_buffer.getvalue(),
                file_name=f"preview_{row[col_nama]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        if st.button("🔄 Generate Semua Surat"):
            output_zip = BytesIO()
            with zipfile.ZipFile(output_zip, "w") as zf:
                for _, row in df.iterrows():
                    doc = Document(uploaded_template)

                    for p in doc.paragraphs:
                        for run in p.runs:
                            if "{{nama}}" in run.text:
                                run.text = run.text.replace("{{nama}}", str(row[col_nama]))

                    for p in doc.paragraphs:
                        if "{{link}}" in p.text:
                            parts = p.text.split("{{link}}")
                            p.clear()
                            if parts[0]: p.add_run(parts[0])
                            add_hyperlink(p, str(row[col_link]), str(row[col_link]))
                            if len(parts) > 1: p.add_run(parts[1])

                    for p in doc.paragraphs:
                        for run in p.runs:
                            run.font.name = "Arial"
                            run.font.size = Pt(12)

                    filename = f"{str(row[col_nama]).replace('/', '-')}.docx"
                    buffer = BytesIO()
                    doc.save(buffer)
                    zf.writestr(filename, buffer.getvalue())

            st.success("✅ Semua surat berhasil dibuat!")
            st.download_button(
                label="📥 Download ZIP Semua Surat",
                data=output_zip.getvalue(),
                file_name="surat_dinamis_output.zip",
                mime="application/zip"
            )
    else:
        st.warning("Excel harus memiliki setidaknya 2 kolom untuk dipilih.")