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

def render_docx_preview_visual(doc):
    st.subheader("📖 Pratinjau Surat (Visual dan Rapi)")
    style = """
    <style>
        .docx-preview {
            background: #fff;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-family: Arial, sans-serif;
            font-size: 15px;
            line-height: 1.6;
            text-align: justify;
            max-height: 400px;
            overflow-y: auto;
        }
        .docx-preview p {
            margin-bottom: 1em;
        }
        .docx-preview a {
            color: #1a0dab;
            text-decoration: underline;
        }
        .docx-preview strong {
            font-weight: bold;
        }
        .docx-preview em {
            font-style: italic;
        }
    </style>
    """
    html = '<div class="docx-preview">'
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        run_html = ""
        for run in p.runs:
            text = run.text.replace("\n", "<br>")
            if run.bold and run.italic:
                run_html += f"<strong><em>{text}</em></strong>"
            elif run.bold:
                run_html += f"<strong>{text}</strong>"
            elif run.italic:
                run_html += f"<em>{text}</em>"
            else:
                run_html += text
        html += f"<p>{run_html}</p>"
    html += "</div>"
    st.markdown(style + html, unsafe_allow_html=True)

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
            status_text.text(f"Memproses surat ke-{idx + 1} dari {total}...")

    output_zip.seek(0)

    if "generate_log" not in st.session_state:
        st.session_state.generate_log = []
    st.session_state.generate_log.extend(log)

    st.session_state.template_count = 1
    st.session_state.last_data_rows = total

    return output_zip, log

SESSION_TIMEOUT = timedelta(minutes=15)

def check_session_timeout():
    if "last_active" in st.session_state:
        if datetime.now() - st.session_state.last_active > SESSION_TIMEOUT:
            st.session_state.clear()
            st.experimental_rerun()
    st.session_state.last_active = datetime.now()

def page_generate():
    st.title("🚀 Generate Surat Massal")

    template_file = st.file_uploader("Upload Template Word (.docx) — Drag & Drop atau klik", type="docx", accept_multiple_files=False)
    data_file = st.file_uploader("Upload Data Excel (.xlsx) — Drag & Drop atau klik", type="xlsx", accept_multiple_files=False)

    if template_file and data_file:
        try:
            df = pd.read_excel(data_file)
            st.success(f"Data Excel berhasil diupload dengan {len(df)} baris")
            st.dataframe(df)

            col_name = st.selectbox("Pilih kolom Nama", df.columns)
            col_link = st.selectbox("Pilih kolom Link", df.columns)

            search_name = st.text_input("Cari Nama (ketik untuk filter)", "")
            filtered_names = df[df[col_name].astype(str).str.contains(search_name, case=False, na=False)][col_name].unique()
            selected_name = st.selectbox("Pilih Nama untuk Preview", filtered_names)

            st.session_state.df = df
            st.session_state.col_name = col_name
            st.session_state.col_link = col_link
            st.session_state.selected_name = selected_name
            st.session_state.template_file = template_file

        except Exception as e:
            st.error(f"Gagal membaca file Excel: {e}")

        if 'show_preview' not in st.session_state:
            st.session_state.show_preview = True

        toggle = st.button("Sembunyikan Preview" if st.session_state.show_preview else "Tampilkan Preview")
        if toggle:
            st.session_state.show_preview = not st.session_state.show_preview

        if st.session_state.show_preview and selected_name:
            row = df[df[col_name] == selected_name].iloc[0]
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

            render_docx_preview_visual(doc)

            preview_buf = BytesIO()
            doc.save(preview_buf)
            preview_buf.seek(0)

            st.download_button(
                label=f"⬇️ Download Preview Surat ({row[col_name]})",
                data=preview_buf.getvalue(),
                file_name=f"preview_{row[col_name]}.docx",
            )

        if st.button("Generate Semua Surat"):
            with st.spinner("Sedang memproses surat..."):
                zip_file, log = generate_letters_with_progress(template_file, df, col_name, col_link)
            st.success("✅ Proses generate selesai!")
            st.toast("Surat berhasil digenerate!", icon="✅")
            st.download_button("Download Semua Surat (ZIP)", zip_file.getvalue(), file_name="surat_massal.zip")
            with st.expander("Lihat Log Generate"):
                st.dataframe(pd.DataFrame(log))

    else:
        st.info("Silakan upload template dan data Excel terlebih dahulu.")

def page_home():
    st.title("🏠 Dashboard")
    st.markdown(f"Selamat datang, **{st.session_state.username}**!")

    with st.expander("Tips Cepat", expanded=True):
        st.write(
            """
            1. Upload template dan data Excel di halaman **Generate Surat**.
            2. Pilih kolom nama dan link sesuai data.
            3. Klik **Generate Semua Surat** dan tunggu hingga selesai.
            4. Unduh file ZIP berisi surat-surat yang sudah jadi.
            """
        )
    st.markdown("---")

    generate_log = st.session_state.get("generate_log", [])

    total_surat = len(generate_log)
    berhasil = sum(1 for item in generate_log if item["Status"].startswith("✅"))
    gagal = total_surat - berhasil

    template_tersedia = st.session_state.get("template_count", 1)
    data_peserta_terakhir = st.session_state.get("last_data_rows", 0)

    statistik_data = {
        "Statistik": [
            "Total Surat Dibuat",
            "Surat Berhasil",
            "Surat Gagal",
            "Template Tersedia",
            "Data Peserta Terakhir",
        ],
        "Jumlah": [
            total_surat,
            berhasil,
            gagal,
            template_tersedia,
            data_peserta_terakhir,
        ],
    }
    df_statistik = pd.DataFrame(statistik_data)
    st.markdown("### Statistik Singkat")
    st.table(df_statistik)

    st.markdown("---")

    st.markdown("### Statistik Surat Berhasil vs Gagal")
    fig, ax = plt.subplots()
    ax.bar(["Berhasil", "Gagal"], [berhasil, gagal], color=["green", "red"])
    ax.set_ylabel("Jumlah Surat")
    ax.set_title("Perbandingan Surat Berhasil dan Gagal")
    st.pyplot(fig)

    st.markdown("---")

    st.markdown("### Persentase Surat")
    fig2, ax2 = plt.subplots()
    if total_surat > 0:
        ax2.pie(
            [berhasil, gagal],
            labels=["Berhasil", "Gagal"],
            autopct="%1.1f%%",
            colors=["green", "red"],
            startangle=90,
            wedgeprops={"edgecolor": "black"},
        )
        ax2.axis("equal")
        st.pyplot(fig2)
    else:
        st.write("Belum ada data surat untuk ditampilkan.")

    st.markdown("---")

    st.markdown("### Aktivitas Terakhir")
    aktivitas = []
    for item in reversed(generate_log[-5:]):
        aktivitas.append({"Aktivitas": f"Generate surat untuk {item['Nama']}", "Status": item["Status"]})
    if aktivitas:
        df_aktivitas = pd.DataFrame(aktivitas)
        st.table(df_aktivitas)
    else:
        st.write("Belum ada aktivitas generate surat.")

    st.markdown("---")

    st.markdown("**Versi Aplikasi:** 1.0.0")
    st.markdown("⚙️ *Tidak ada pemeliharaan sistem saat ini.*")

def show_login():
    st.set_page_config(page_title="Generator Surat Hyperlink", layout="centered")
    st.title("📬 Selamat Datang di Aplikasi Surat Massal PMT")
    st.markdown("Silakan login untuk menggunakan aplikasi ini.")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            if username == "admin" and password == "surat123":
                st.session_state.login_state = True
                st.session_state.username = username
                st.session_state.logout_message = False
                st.experimental_rerun()
            else:
                st.error("Username atau password salah.")

def show_main_app():
    check_session_timeout()
    st.sidebar.success(f"Login sebagai: {st.session_state.username}")
    if st.sidebar.button("Logout"):
        st.session_state.logout_message = True
        st.session_state.login_state = False
        st.session_state.username = ""
        st.experimental_rerun()

    st.sidebar.title("Menu")
    page = st.sidebar.radio("Navigasi", ["Dashboard", "Generate Surat"])

    if page == "Dashboard":
        page_home()
    elif page == "Generate Surat":
        page_generate()

if "login_state" not in st.session_state:
    st.session_state.login_state = False

if st.session_state.get("logout_message", False):
    st.set_page_config(page_title="Sampai Jumpa!", layout="centered")
    st.title("👋 Terima Kasih!")
    st.markdown("Terima kasih telah menggunakan aplikasi ini.\n\n**See you!**")
    if st.button("🔐 Kembali ke Halaman Login"):
        st.session_state.logout_message = False
        st.experimental_rerun()
elif st.session_state.login_state:
    show_main_app()
else:
    show_login()
