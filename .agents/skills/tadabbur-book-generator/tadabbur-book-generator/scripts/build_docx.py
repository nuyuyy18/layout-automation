import sys, os, docx
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def add_callout_box(doc, text_content, title="[ Dialog Mindful ]"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Cm(10.8)
    set_cell_background(cell, "FAFAFA")
    set_cell_margins(cell, top=160, bottom=160, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="000000"/>'
        f'<w:top w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:bottom w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p_title = cell.paragraphs[0]
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    run_t = p_title.add_run(title)
    run_t.font.name = "Georgia"
    run_t.font.size = Pt(12)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(0, 0, 0)
    
    for line in text_content.strip().split("\n"):
        if not line.strip():
            continue
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.2
        run = p.add_run(line.strip())
        run.font.name = "Georgia"
        run.font.size = Pt(11)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

def create_a5_tadabbur_docx(chapter_data, output_path):
    """
    chapter_data is a dict containing 7-pillar information.
    """
    doc = Document()
    
    section = doc.sections[0]
    section.page_width = Cm(14.8)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(1.8)
    
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Georgia'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    COLOR_BLACK = RGBColor(0, 0, 0)
    
    # H1 Bab (Page Break Before)
    p_chap = doc.add_paragraph()
    p_chap.paragraph_format.space_before = Pt(18)
    p_chap.paragraph_format.space_after = Pt(2)
    run_chap = p_chap.add_run(chapter_data.get('chapter_num', 'BAB 1'))
    run_chap.font.size = Pt(15)
    run_chap.font.bold = True
    run_chap.font.color.rgb = COLOR_BLACK
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(12)
    p_title.paragraph_format.line_spacing = 1.15
    run_title = p_title.add_run(chapter_data.get('chapter_title', ''))
    run_title.font.size = Pt(19)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_BLACK
    
    # 1. Ayat & Doa
    p_sec0 = doc.add_paragraph()
    p_sec0.paragraph_format.space_before = Pt(8)
    p_sec0.paragraph_format.space_after = Pt(4)
    run_sec0 = p_sec0.add_run("[ Ayat Pijakan & Doa Hati ]")
    run_sec0.font.bold = True
    run_sec0.font.size = Pt(13)
    run_sec0.font.color.rgb = COLOR_BLACK
    
    if chapter_data.get('ayat_ar'):
        p_ar = doc.add_paragraph()
        p_ar.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_ar = p_ar.add_run(chapter_data['ayat_ar'])
        r_ar.font.size = Pt(14)
        r_ar.font.bold = True
        r_ar.font.color.rgb = COLOR_BLACK
        
    if chapter_data.get('ayat_id'):
        p_id = doc.add_paragraph()
        p_id.paragraph_format.space_after = Pt(8)
        p_id.paragraph_format.line_spacing = 1.2
        r_id = p_id.add_run(chapter_data['ayat_id'])
        r_id.font.size = Pt(11)
        r_id.font.italic = True
        r_id.font.color.rgb = COLOR_BLACK

    if chapter_data.get('doa_ar'):
        p_dar = doc.add_paragraph()
        p_dar.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_dar = p_dar.add_run(chapter_data['doa_ar'])
        r_dar.font.size = Pt(13)
        r_dar.font.bold = True
        r_dar.font.color.rgb = COLOR_BLACK
        
    if chapter_data.get('doa_id'):
        p_did = doc.add_paragraph()
        p_did.paragraph_format.space_after = Pt(14)
        p_did.paragraph_format.line_spacing = 1.2
        r_did = p_did.add_run(chapter_data['doa_id'])
        r_did.font.size = Pt(11)
        r_did.font.italic = True
        r_did.font.color.rgb = COLOR_BLACK

    # 2. Catatan Hati
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(10)
    p_sec1.paragraph_format.space_after = Pt(2)
    run_sec1 = p_sec1.add_run("[ Catatan Hati ]")
    run_sec1.font.bold = True
    run_sec1.font.size = Pt(13)
    run_sec1.font.color.rgb = COLOR_BLACK
    
    if chapter_data.get('catatan_hati_title'):
        p_sub1 = doc.add_paragraph()
        p_sub1.paragraph_format.space_after = Pt(2)
        r_sub1 = p_sub1.add_run(chapter_data['catatan_hati_title'])
        r_sub1.font.bold = True
        r_sub1.font.size = Pt(14)
        r_sub1.font.color.rgb = COLOR_BLACK
        
    if chapter_data.get('catatan_hati_subtitle'):
        p_sub_sub = doc.add_paragraph()
        p_sub_sub.paragraph_format.space_after = Pt(10)
        r_sss = p_sub_sub.add_run(chapter_data['catatan_hati_subtitle'])
        r_sss.font.italic = True
        r_sss.font.size = Pt(11)
        r_sss.font.color.rgb = COLOR_BLACK
        
    for text in chapter_data.get('catatan_hati_paras', []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.first_line_indent = Cm(0.5)
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_BLACK

    # 3. Sains Jiwa
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(12)
    p_sec2.paragraph_format.space_after = Pt(2)
    run_sec2 = p_sec2.add_run(f"[ {chapter_data.get('sains_label', 'Sains Jiwa')} ]")
    run_sec2.font.bold = True
    run_sec2.font.size = Pt(13)
    run_sec2.font.color.rgb = COLOR_BLACK
    
    if chapter_data.get('sains_title'):
        p_sub2 = doc.add_paragraph()
        p_sub2.paragraph_format.space_after = Pt(8)
        r_sub2 = p_sub2.add_run(chapter_data['sains_title'])
        r_sub2.font.bold = True
        r_sub2.font.size = Pt(14)
        r_sub2.font.color.rgb = COLOR_BLACK
        
    for text in chapter_data.get('sains_paras', []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.first_line_indent = Cm(0.5)
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_BLACK

    # 4. Dialog Mindful
    if chapter_data.get('mindful_text'):
        add_callout_box(doc, chapter_data['mindful_text'], title=chapter_data.get('mindful_title', '[ Dialog Mindful ]'))

    # 5. Amal Tracker (1x)
    p_sec4 = doc.add_paragraph()
    p_sec4.paragraph_format.space_before = Pt(10)
    p_sec4.paragraph_format.space_after = Pt(2)
    run_sec4 = p_sec4.add_run("[ Amal Tracker ]")
    run_sec4.font.bold = True
    run_sec4.font.size = Pt(13)
    run_sec4.font.color.rgb = COLOR_BLACK
    
    p_sub4 = doc.add_paragraph()
    p_sub4.paragraph_format.space_after = Pt(6)
    r_sub4 = p_sub4.add_run("Langkah Aksi Hari Ini")
    r_sub4.font.bold = True
    r_sub4.font.size = Pt(13.5)
    r_sub4.font.color.rgb = COLOR_BLACK
    
    for label, desc in chapter_data.get('checklist_items', []):
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_after = Pt(4)
        p_c.paragraph_format.line_spacing = 1.2
        r_box = p_c.add_run("[  ]  ")
        r_box.font.bold = True
        r_box.font.size = Pt(12)
        r_box.font.color.rgb = COLOR_BLACK
        r_lbl = p_c.add_run(f"{label}: ")
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(11.5)
        r_lbl.font.color.rgb = COLOR_BLACK
        r_desc = p_c.add_run(desc)
        r_desc.font.size = Pt(11.5)
        r_desc.font.color.rgb = COLOR_BLACK

    # 6. Pojok Riset
    p_sec5 = doc.add_paragraph()
    p_sec5.paragraph_format.space_before = Pt(10)
    p_sec5.paragraph_format.space_after = Pt(2)
    run_sec5 = p_sec5.add_run("[ Pojok Riset ]")
    run_sec5.font.bold = True
    run_sec5.font.size = Pt(13)
    run_sec5.font.color.rgb = COLOR_BLACK
    
    if chapter_data.get('riset_title'):
        p_sub5 = doc.add_paragraph()
        p_sub5.paragraph_format.space_after = Pt(4)
        r_sub5 = p_sub5.add_run(chapter_data['riset_title'])
        r_sub5.font.bold = True
        r_sub5.font.size = Pt(13)
        r_sub5.font.color.rgb = COLOR_BLACK
        
    if chapter_data.get('riset_text'):
        p_riset = doc.add_paragraph()
        p_riset.paragraph_format.space_after = Pt(12)
        p_riset.paragraph_format.line_spacing = 1.2
        r_r = p_riset.add_run(chapter_data['riset_text'])
        r_r.font.size = Pt(11)
        r_r.font.italic = True
        r_r.font.color.rgb = COLOR_BLACK

    # 7. Kolom Interaktif & Quote
    p_sec6 = doc.add_paragraph()
    p_sec6.paragraph_format.space_before = Pt(10)
    p_sec6.paragraph_format.space_after = Pt(2)
    run_sec6 = p_sec6.add_run("[ Kolom Interaktif ]")
    run_sec6.font.bold = True
    run_sec6.font.size = Pt(13)
    run_sec6.font.color.rgb = COLOR_BLACK
    
    if chapter_data.get('interaktif_title'):
        p_sub6 = doc.add_paragraph()
        p_sub6.paragraph_format.space_after = Pt(6)
        r_sub6 = p_sub6.add_run(chapter_data['interaktif_title'])
        r_sub6.font.bold = True
        r_sub6.font.size = Pt(13.5)
        r_sub6.font.color.rgb = COLOR_BLACK
        
    for q in chapter_data.get('questions', []):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.2
        run = p.add_run(q)
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_BLACK
        if "." in q[:2] or '"' in q[:1]:
            run.font.bold = True

    if chapter_data.get('quote'):
        doc.add_paragraph().paragraph_format.space_after = Pt(8)
        p_qtitle = doc.add_paragraph()
        p_qtitle.paragraph_format.space_before = Pt(8)
        p_qtitle.paragraph_format.space_after = Pt(2)
        r_qt = p_qtitle.add_run("[ Quote of the Day ]")
        r_qt.font.bold = True
        r_qt.font.size = Pt(13)
        r_qt.font.color.rgb = COLOR_BLACK
        
        p_quote = doc.add_paragraph()
        p_quote.paragraph_format.space_after = Pt(14)
        p_quote.paragraph_format.line_spacing = 1.25
        r_q = p_quote.add_run(f'"{chapter_data["quote"].strip()}"')
        r_q.font.size = Pt(12)
        r_q.font.italic = True
        r_q.font.bold = True
        r_q.font.color.rgb = COLOR_BLACK
        
    doc.save(output_path)
    print("SUCCESS: A5 DOCX generated at", output_path)
