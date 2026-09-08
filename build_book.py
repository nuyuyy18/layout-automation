import os
import json
import re
import docx
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Palet Warna & Tipografi Baku
C_BLACK = RGBColor(0, 0, 0)
C_DARK_GREEN = RGBColor(27, 94, 32)
C_MUTED = RGBColor(80, 80, 80)
C_LIGHT = RGBColor(120, 120, 120)

def clean_all(text):
    if not text:
        return ""
    # Hapus tanda pisah em-dash dan en-dash secara tuntas
    t = text.replace("—", ", ").replace("–", " ")
    # Hapus tulisan tere liye (case-insensitive)
    t = re.sub(r'tere\s*liye,?\s*', '', t, flags=re.IGNORECASE)
    t = t.replace(", ,", ",").replace(" ,", ",")
    while "  " in t:
        t = t.replace("  ", " ")
    return t.strip()

def set_cell(cell, fill_hex="FFFFFF", top=80, bottom=80, left=100, right=100, borders=None):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
    ))
    if borders:
        b_xml = f'<w:tcBorders {nsdecls("w")}>'
        for side in ["top", "left", "bottom", "right"]:
            b = borders.get(side)
            b_xml += f'<w:{side} w:val="{b.get("val", "single")}" w:sz="{b.get("sz", "4")}" w:space="0" w:color="{b.get("color", "000000")}"/>' if b else f'<w:{side} w:val="none"/>'
        b_xml += '</w:tcBorders>'
        tcPr.append(parse_xml(b_xml))

def add_p(doc, text="", bold=False, italic=False, size=11.5, space_after=6, space_before=0, line_spacing=1.25, indent=0.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY, color=C_BLACK):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent > 0:
        p.paragraph_format.first_line_indent = Cm(indent)
    if text:
        cleaned = clean_all(text)
        r = p.add_run(cleaned)
        r.font.name = "Garamond"
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return p

def add_box(doc, title, text, bg="FBFBF9", border_color="1B5E20", italic=True):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c = tbl.cell(0, 0)
    c.width = Cm(10.8)
    set_cell(c, fill_hex=bg, top=140, bottom=140, left=160, right=160,
             borders={"left": {"val": "single", "sz": "16", "color": border_color},
                      "top": {"val": "single", "sz": "4", "color": "E0E0E0"},
                      "bottom": {"val": "single", "sz": "4", "color": "E0E0E0"},
                      "right": {"val": "single", "sz": "4", "color": "E0E0E0"}})
    p0 = c.paragraphs[0]
    p0.paragraph_format.space_after = Pt(4)
    r0 = p0.add_run(clean_all(title))
    r0.font.name = "Garamond"
    r0.font.size = Pt(10.5)
    r0.font.bold = True
    r0.font.color.rgb = C_DARK_GREEN
    
    for line in text.strip().split("\n"):
        cleaned_line = clean_all(line)
        if cleaned_line:
            p = c.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.2
            r = p.add_run(cleaned_line)
            r.font.name = "Garamond"
            r.font.size = Pt(10.5)
            r.font.italic = italic
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_arabic_box(doc, ayat_ar, ayat_id, ayat_ref, doa_ar=None, doa_id=None, doa_ref=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c = tbl.cell(0, 0)
    c.width = Cm(10.8)
    set_cell(c, fill_hex="FBFBF9", top=140, bottom=140, left=160, right=160,
             borders={"left": {"val": "single", "sz": "16", "color": "1B5E20"},
                      "top": {"val": "single", "sz": "4", "color": "CCCCCC"},
                      "bottom": {"val": "single", "sz": "4", "color": "CCCCCC"},
                      "right": {"val": "single", "sz": "4", "color": "CCCCCC"}})
    p0 = c.paragraphs[0]
    p0.paragraph_format.space_after = Pt(4)
    r0 = p0.add_run("📖  AYAT PIJAKAN & DOA HATI")
    r0.font.name = "Garamond"
    r0.font.size = Pt(10.5)
    r0.font.bold = True
    r0.font.color.rgb = C_DARK_GREEN
    
    for ar, trans, ref in [(ayat_ar, ayat_id, ayat_ref), (doa_ar, doa_id, doa_ref)]:
        if ar:
            p_ar = c.add_paragraph()
            p_ar.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_ar.paragraph_format.line_spacing = 1.6
            r_ar = p_ar.add_run(ar.strip())
            r_ar.font.name = "Traditional Arabic"
            r_ar.font.size = Pt(14)
            r_ar.font.bold = True
        if trans:
            p_tr = c.add_paragraph()
            p_tr.paragraph_format.line_spacing = 1.2
            r_tr = p_tr.add_run(f'"{clean_all(trans)}"')
            r_tr.font.name = "Garamond"
            r_tr.font.size = Pt(10)
            r_tr.font.italic = True
            if ref:
                p_ref = c.add_paragraph()
                p_ref.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r_ref = p_ref.add_run(f"({clean_all(ref)})")
                r_ref.font.name = "Garamond"
                r_ref.font.size = Pt(9)
                r_ref.font.bold = True
                r_ref.font.color.rgb = C_DARK_GREEN
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def generate_book(json_path="book_content.json", out_docx="result/Ketika_Jiwa_Pulang_ke_Sunyi.docx"):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(14.8), Cm(21.0)
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(2.0)
    sec.left_margin, sec.right_margin = Cm(2.2), Cm(1.8)
    sec.different_first_page_header_footer = True
    
    # Nomor Halaman Footer Bersih (Hanya Angka, Tanpa Tanda Hubung)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>'))
    
    fm = data["front_matter"]
    
    # 1. Cover
    doc.add_paragraph().paragraph_format.space_before = Pt(40)
    add_p(doc, "❖   ❖   ❖", size=14, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=12)
    add_p(doc, fm["title"], bold=True, size=22, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=8)
    add_p(doc, fm["subtitle"], italic=True, size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=28)
    add_p(doc, fm["compiler"], bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=6)
    for insp in fm["inspirations"]:
        add_p(doc, insp, italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=2)
    doc.add_paragraph().paragraph_format.space_before = Pt(56)
    add_p(doc, fm["edition"], size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=2)
    add_p(doc, f"Yogyakarta, {fm['year']}", size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_LIGHT, indent=0)

    # 2. Kolofon & Epigraph
    doc.add_page_break()
    doc.add_paragraph().paragraph_format.space_before = Pt(36)
    add_p(doc, fm["title"], bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4)
    add_p(doc, fm["subtitle"], italic=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=24)
    add_p(doc, "❖", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=36)
    add_p(doc, "HAK CIPTA & KETERANGAN PENERBITAN", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=8)
    add_p(doc, "Buku ini disusun sebagai materi refleksi spiritual dan psikologi islami berbasis riset ilmiah internasional.", italic=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=4)
    add_p(doc, "Format Buku: Ukuran Standar A5 (14.8 cm x 21.0 cm) Monokrom Bersih.", size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=36)
    add_p(doc, f"Cetakan Pertama: September {fm['year']}", size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0)
    
    doc.add_page_break()
    doc.add_paragraph().paragraph_format.space_before = Pt(80)
    author_text = clean_all(fm["epigraph"].get("author", "Catatan Perjalanan Batin"))
    add_box(doc, "KUTIPAN PEMBUKA", fm["epigraph"]["quote"] + f"\n\n({author_text})", bg="FBFBF9")

    # 3. Prolog & TOC
    doc.add_page_break()
    add_p(doc, "PROLOG", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
    add_p(doc, fm["prologue"]["title"], bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=16)
    add_p(doc, "❖", size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_LIGHT, indent=0, space_after=16)
    for p in fm["prologue"]["paragraphs"]:
        add_p(doc, p)
        
    doc.add_page_break()
    add_p(doc, "DAFTAR ISI LENGKAP", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
    add_p(doc, "Peta Perjalanan Jiwa Menuju Ketenangan Sejati", italic=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=18)
    for part in data["parts"]:
        add_p(doc, f"{part['num']}: {part['title']}", bold=True, size=10.5, color=C_DARK_GREEN, indent=0, space_after=2)
        for chap in part["chapters"]:
            add_p(doc, f"  {chap['num']}: {chap['title']}", size=9.5, indent=0.4, space_after=2)
    add_p(doc, "Epilog: Sebuah Surat untuk Jiwa yang Berjuang", bold=True, size=10, indent=0, space_after=2)
    add_p(doc, "Jurnal Interaktif, Catatan Notebook, & Riset Ilmiah", bold=True, size=10, indent=0, space_after=2)

    # 4. Bagian & Bab-Bab
    for part in data["parts"]:
        doc.add_page_break()
        doc.add_paragraph().paragraph_format.space_before = Pt(72)
        add_p(doc, part["num"].upper(), bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=8)
        add_p(doc, part["title"], bold=True, size=18, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=14)
        add_p(doc, part["subtitle"], italic=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=20)
        add_p(doc, f"Inspirasi Kajian: {part['source']}", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_LIGHT, indent=0)
        
        for chap in part["chapters"]:
            doc.add_page_break()
            add_p(doc, chap["num"].upper(), bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
            add_p(doc, chap["title"], bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=6)
            if chap.get("subtitle"):
                add_p(doc, chap["subtitle"], italic=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=14)
            add_p(doc, "❖", size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_LIGHT, indent=0, space_after=16)
            
            add_arabic_box(doc, chap.get("ayat_ar"), chap.get("ayat_id"), chap.get("ayat_ref"),
                           chap.get("doa_ar"), chap.get("doa_id"), chap.get("doa_ref"))
            
            add_p(doc, "[ Catatan Hati: Refleksi Perjalanan Batin ]", bold=True, size=12, color=C_DARK_GREEN, indent=0, space_after=4)
            for p in chap["catatan_hati"]:
                add_p(doc, p)
                
            add_p(doc, f"[ {chap['sains_label']}: {chap['sains_title']} ]", bold=True, size=12, color=C_DARK_GREEN, indent=0, space_after=4)
            for p in chap["sains_paras"]:
                add_p(doc, p, size=11, space_after=5)
                
            add_box(doc, f"🌿  {chap['mindful_title']}", chap["mindful_text"], bg="F8F9FA")
            add_box(doc, f"KUTIPAN HIKMAH ({chap['num']})", f'"{chap["quote"]}"', bg="F5F5F0", italic=True)

    # 5. Epilog
    doc.add_page_break()
    add_p(doc, "EPILOG", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
    add_p(doc, data["epilogue"]["title"], bold=True, size=15, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=16)
    add_p(doc, "❖", size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_LIGHT, indent=0, space_after=16)
    for p in data["epilogue"]["paragraphs"]:
        add_p(doc, p)

    # 6. Ceklis Interaktif (1 Halaman per Kategori)
    for cat_title, items in data["checklist_data"]:
        doc.add_page_break()
        add_p(doc, "JURNAL INTERAKTIF: SELF-REFLECTION TRACKER", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4)
        add_p(doc, cat_title, bold=True, size=13.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=12)
        
        tbl = doc.add_table(rows=len(items)+1, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        widths = [Cm(1.4), Cm(4.8), Cm(4.6)]
        for i, h in enumerate(["Ceklis", "Langkah Aksi / Amal", "Target & Refleksi Batin"]):
            c = tbl.rows[0].cells[i]
            c.width = widths[i]
            set_cell(c, fill_hex="1B5E20", top=100, bottom=100, left=100, right=100)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            r.font.name, r.font.size, r.font.bold = "Garamond", Pt(9.5), True
            r.font.color.rgb = RGBColor(255, 255, 255)
            
        for row_idx, (chk, amal, tgt) in enumerate(items, start=1):
            bg = "F9F9F9" if row_idx % 2 == 0 else "FFFFFF"
            for ci, txt in enumerate([chk, amal, tgt]):
                c = tbl.rows[row_idx].cells[ci]
                c.width = widths[ci]
                set_cell(c, fill_hex=bg, top=80, bottom=80, left=90, right=90,
                         borders={"top": {"sz": "4", "color": "E0E0E0"}, "bottom": {"sz": "4", "color": "E0E0E0"},
                                  "left": {"sz": "4", "color": "E0E0E0"}, "right": {"sz": "4", "color": "E0E0E0"}})
                p = c.paragraphs[0]
                if ci == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(txt)
                    r.font.name, r.font.size, r.font.bold = "Garamond", Pt(11), True
                elif ci == 1:
                    r = p.add_run(clean_all(txt))
                    r.font.name, r.font.size, r.font.bold = "Garamond", Pt(10), True
                else:
                    r = p.add_run(clean_all(txt))
                    r.font.name, r.font.size = "Garamond", Pt(9.5)
                    r.font.color.rgb = C_MUTED

    # 7. Notebook Curhat Batin (2 Halaman)
    for nb in data["notebook_data"]:
        doc.add_page_break()
        add_p(doc, f"LEMBAR CURHAT BATIN & PROGRESS MINGGUAN (BAGIAN {nb['sheet_num']})", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=4)
        add_p(doc, "Ruang hening ini disediakan khusus untuk tulisan tanganmu di hadapan Allah:", italic=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=10)
        
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        c = tbl.cell(0, 0)
        c.width = Cm(10.8)
        set_cell(c, fill_hex="FBFBF9", top=140, bottom=140, left=160, right=160,
                 borders={"left": {"val": "single", "sz": "16", "color": "1B5E20"},
                          "top": {"val": "single", "sz": "10", "color": "1B5E20"},
                          "bottom": {"val": "single", "sz": "10", "color": "1B5E20"},
                          "right": {"val": "single", "sz": "6", "color": "CCCCCC"}})
        p0 = c.paragraphs[0]
        p0.paragraph_format.space_after = Pt(4)
        r0 = p0.add_run(f"✍️  {clean_all(nb['title']).upper()}")
        r0.font.name, r0.font.size, r0.font.bold, r0.font.color.rgb = "Garamond", Pt(11), True, C_DARK_GREEN
        
        p_q = c.add_paragraph()
        p_q.paragraph_format.space_after = Pt(8)
        r_q = p_q.add_run(f'"{clean_all(nb["quote"])}"')
        r_q.font.name, r_q.font.size, r_q.font.italic, r_q.font.color.rgb = "Garamond", Pt(9.5), True, C_MUTED
        
        for prompt in nb["prompts"]:
            p = c.add_paragraph()
            p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(1), Pt(3)
            p.paragraph_format.line_spacing = 1.25
            r = p.add_run(clean_all(prompt))
            r.font.name, r.font.size = "Garamond", Pt(9.5)

    # 8. Riset Ilmiah & Integrasi QR
    doc.add_page_break()
    add_p(doc, "POJOK RISET & VALIDASI ILMIAH", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
    add_p(doc, "Daftar Sitasi Jurnal Peer-Reviewed dengan DOI Aktif", italic=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_MUTED, indent=0, space_after=16)
    for idx, r in enumerate(data["research_data"], start=1):
        add_p(doc, f"[{idx}] {clean_all(r['citation'])}", bold=True, size=9.5, space_after=1, indent=0)
        add_p(doc, f"DOI Resmi: {r['doi']}", size=8.5, color=C_DARK_GREEN, space_after=1, indent=0)
        add_p(doc, f"Signifikansi Riset: {clean_all(r['notes'])}", italic=True, size=9, color=C_MUTED, space_after=8, indent=0)

    res_dir = os.path.dirname(out_docx)
    for i, qr in enumerate(data["digital_integration"]["qr_items"]):
        doc.add_page_break() # Setiap QR Code menempati halaman baru agar teks dan gambar selalu utuh di satu halaman
        if i == 0:
            add_p(doc, "INTEGRASI DIGITAL & KODE QR PENDUKUNG", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=4)
            add_p(doc, "Pindai QR Code untuk konsultasi AI Assistant atau menyaksikan video sumber:", size=10, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, space_after=14)
        else:
            add_p(doc, "INTEGRASI DIGITAL & KODE QR PENDUKUNG (LANJUTAN)", bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=C_DARK_GREEN, indent=0, space_after=12)
            
        add_box(doc, qr["title"], f"{qr['desc']}\n\nTautan Resmi: {qr['url']}", bg="FAFAFA", italic=False)
        img_p = os.path.join(res_dir, qr["img_file"])
        if os.path.exists(img_p):
            p_img = doc.paragraphs[-1]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(8)
            p_img.add_run().add_picture(img_p, width=Inches(1.5))
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

    os.makedirs(os.path.dirname(out_docx), exist_ok=True)
    doc.save(out_docx)
    print("SUCCESS: Saved", out_docx)

if __name__ == "__main__":
    generate_book()
