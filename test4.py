import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from io import BytesIO
import zipfile
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import unicodedata, string, re

# --- KONFIGURASI STREAMLIT ---
st.set_page_config(page_title="Surat Massal PMT", layout="centered")

# --- KAMUS BAHASA (sama seperti punya kamu) ---
LANGUAGES = { ... }  # <-- pakai yang sudah ada di kode kamu

def t(key):
    return LANGUAGES.get(st.session_state.lang, LANGUAGES["id"]).get(key, key)

# --- HELPER FILENAME ---
def safe_filename_base(name: str):
    valid = f"-_.() {string.ascii_letters}{string.digits}"
    norm = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    cleaned = "".join(c for c in norm if c in valid).strip()
    return cleaned or "surat"

# --- HAPUS PARAGRAF ---
def _clear_paragraph(p):
    for r in p.runs:
        r._r.getparent().remove(r._r)

# --- GANTI PLACEHOLDER LINK ---
def replace_placeholder_with_hyperlink(p, text_before, url, text_after):
    _clear_paragraph(p)
    if text_before:
        rb = p.add_run(text_before)
        rb.font.name = "Arial"
        rb.font.size = Pt(12)

    if url:
        r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
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
        text_elem.text = str(url)
        new_run.append(text_elem)
        hyperlink.append(new_run)
        p._p.append(hyperlink)
    else:
        rp = p.add_run("(tautan tidak tersedia)")
        rp.font.name = "Arial"
        rp.font.size = Pt(12)

    if text_after:
        ra = p.add_run(text_after)
        ra.font.name = "Arial"
        ra.font.size = Pt(12)

    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# --- NORMALISASI LINK ---
def normalize_url(u: str):
    u = ("" if pd.isna(u) else str(u)).strip()
    if not u:
        return ""
    if not (u.startswith("http://") or u.startswith("https://")):
        u = "https://" + u
    return u

# --- GENERATE LETTERS ---
def generate_letters_with_progress(template_file, df, col_name, col_link):
    df[col_link] = df[col_link].map(normalize_url)

    output_zip = BytesIO()
    log = []
    pad_width = st.session_state.get("pad_width", len(str(len(df))))
    used_names = set()

    with zipfile.ZipFile(output_zip, "w") as zf:
        total = len(df)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, (idx, row) in enumerate(df.iterrows(), start=1):
            try:
                tpl = DocxTemplate(template_file)
                tpl.render({
                    "nama_penyelenggara": row[col_name],
                    "short_link": "[short_link]",
                    "no": int(row["No"]),  # nomor urut
                })
                temp_buf = BytesIO()
                tpl.save(temp_buf)
                temp_buf.seek(0)

                doc = Document(temp_buf)
                pattern = re.escape("[short_link]")
                for p in doc.paragraphs:
                    if re.search(pattern, p.text):
                        parts = p.text.split("[short_link]")
                        before = parts[0] if len(parts) > 0 else ""
                        after = parts[1] if len(parts) > 1 else ""
                        replace_placeholder_with_hyperlink(p, before, row[col_link], after)

                # Styling font kecuali hyperlink
                for p in doc.paragraphs:
                    if p._p.xpath(".//w:hyperlink"):
                        continue
                    for run in p.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(12)

                final_buf = BytesIO()
                doc.save(final_buf)

                # Prefix nomor di nama file
                base = safe_filename_base(row[col_name])
                prefix = f"{int(row['No']):0{pad_width}d}"
                file_name = f"{prefix} - {base}.docx"
                n = 1
                candidate = file_name
                while candidate.lower() in used_names:
                    n += 1
                    candidate = f"{prefix} - {base} ({n}).docx"
                used_names.add(candidate.lower())

                zf.writestr(candidate, final_buf.getvalue())
                log.append({"Nama": row[col_name], "Status": "✅ Berhasil"})
            except Exception as e:
                log.append({"Nama": row.get(col_name, '(unknown)'), "Status": f"❌ Gagal: {e}"})

            progress = int(i / total * 100)
            progress_bar.progress(progress)
            status_text.text(f"{t('processing_letters')} {i} / {total}")

    output_zip.seek(0)
    st.session_state.generate_log = st.session_state.get("generate_log", []) + log
    st.session_state.template_count = 1
    st.session_state.last_data_rows = len(df)
    return output_zip, log

# --- PAGE GENERATE ---
def page_generate():
    st.title(t("generate_title"))
    template_file = st.file_uploader(t("upload_template"), type="docx")
    data_file = st.file_uploader(t("upload_data"), type="xlsx")

    if template_file and data_file:
        try:
            df = pd.read_excel(data_file)

            # Tambah kolom No jika belum ada
            if "No" not in df.columns:
                df.insert(0, "No", range(1, len(df) + 1))
            pad_width = len(str(len(df)))
            st.session_state.pad_width = pad_width

            st.success(f"{len(df)} rows loaded successfully")
            st.dataframe(df)

            col_name = st.selectbox(t("select_name_col"), df.columns)
            col_link = st.selectbox(t("select_link_col"), df.columns)
            search_name = st.text_input(t("search_name"), "")
            filtered_names = df[df[col_name].astype(str).str.contains(search_name, case=False, na=False)][col_name].unique()
            selected_name = st.selectbox(t("select_name_preview"), filtered_names)

            st.session_state.df = df
            st.session_state.col_name = col_name
            st.session_state.col_link = col_link
            st.session_state.selected_name = selected_name
            st.session_state.template_file = template_file

        except Exception as e:
            st.error(f"Failed to read Excel file: {e}")
            return

        if 'show_preview' not in st.session_state:
            st.session_state.show_preview = True

        toggle = st.button(t("hide_preview") if st.session_state.show_preview else t("show_preview"))
        if toggle:
            st.session_state.show_preview = not st.session_state.show_preview

        if st.session_state.show_preview and selected_name:
            row = df[df[col_name] == selected_name].iloc[0]
            tpl = DocxTemplate(template_file)
            tpl.render({
                "nama_penyelenggara": row[col_name],
                "short_link": "[short_link]",
                "no": int(row["No"]),
            })
            temp_buf = BytesIO()
            tpl.save(temp_buf)
            temp_buf.seek(0)

            doc = Document(temp_buf)
            for p in doc.paragraphs:
                if "[short_link]" in p.text:
                    parts = p.text.split("[short_link]")
                    before = parts[0] if len(parts) > 0 else ""
                    after = parts[1] if len(parts) > 1 else ""
                    replace_placeholder_with_hyperlink(p, before, row[col_link], after)

            render_docx_preview_visual(doc)  # pakai fungsi kamu yg sudah ada

            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)
            st.download_button(f"{t('download_preview')} ({row[col_name]})", preview_buf.getvalue(), file_name=f"preview_{row[col_name]}.docx")

        if st.button(t("generate_all")):
            with st.spinner(t("processing_letters")):
                zip_file, log = generate_letters_with_progress(template_file, df, col_name, col_link)
            st.success(t("generate_done"))
            st.download_button(t("download_all_zip"), zip_file.getvalue(), file_name="surat_massal.zip")
            with st.expander(t("view_log")):
                st.dataframe(pd.DataFrame(log))
    else:
        st.info(t("upload_first"))
