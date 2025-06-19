print("THIS IS THE FILE BEING RUN")
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def rgb_color_from_hex(hex_color):
    hex_color = hex_color.lstrip('#')
    return RGBColor(*(int(hex_color[i:i+2], 16) for i in (0, 2, 4)))

def set_strict_page_margins_fixed(section, inches=1):
    section.top_margin = Inches(inches)
    section.bottom_margin = Inches(inches)
    section.left_margin = Inches(inches)
    section.right_margin = Inches(inches)

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def insert_horizontal_line(paragraph, color="auto"):
    p = paragraph._p  # Access the internal lxml element
    pPr = p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)

def add_header_with_fully_flush_left_logo(doc, logo_path, bar_color_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    section = doc.sections[0]
    header = section.header
    section.header_distance = Inches(0.15)
    for para in header.paragraphs:
        p = para._element
        p.getparent().remove(p)
    logo_para = header.add_paragraph()
    logo_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    logo_para.paragraph_format.left_indent = Inches(-0.7)
    logo_para.paragraph_format.space_before = Pt(0)
    logo_para.paragraph_format.space_after = Pt(0)
    run = logo_para.add_run()
    run.add_picture(logo_path, width=Inches(1))
    bar_para = header.add_paragraph()
    bar_para.paragraph_format.space_before = Pt(0)
    bar_para.paragraph_format.space_after = Pt(0)
    bar_para.paragraph_format.left_indent = Inches(-0.7)
    bar_para.paragraph_format.right_indent = Inches(-0.7)
    bar_para.paragraph_format.line_spacing = Pt(1)
    bar_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    border_xml = (
        f'<w:pBorders {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="40" w:color="{bar_color_hex.lstrip("#")}" w:space="0"/>'
        f'</w:pBorders>'
    )
    bar_para._p.get_or_add_pPr().append(parse_xml(border_xml))
    return doc
  
