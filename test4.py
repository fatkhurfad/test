from docxtpl import DocxTemplate
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from io import BytesIO

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")  # 12 pt
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

# === Langkah 1: Render dari Template ===
tpl = DocxTemplate("template_with_placeholders.docx")
context = {
    "nama_penyelenggara": "PT 3D Tech",
    "short_link": "[short_link]"  # gunakan token sementara
}
tpl.render(context)

buffer = BytesIO()
tpl.save(buffer)
buffer.seek(0)

# === Langkah 2: Buka ulang dan ganti [short_link] dengan hyperlink ===
doc = Document(buffer)
for p in doc.paragraphs:
    if "[short_link]" in p.text:
        parts = p.text.split("[short_link]")
        p.clear()
        if parts[0]: p.add_run(parts[0])
        add_hyperlink(p, "https://komdigi.go.id", "https://komdigi.go.id")
        if len(parts) > 1: p.add_run(parts[1])

# === Simpan Hasil Akhir ===
doc.save("surat_final_PT_3D_Tech.docx")
print("✅ Surat berhasil dibuat dengan hyperlink dan header tetap aman.")
