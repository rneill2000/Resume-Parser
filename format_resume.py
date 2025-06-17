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

def insert_horizontal_line(paragraph, hex_color):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    # Remove existing borders if any
    for border in pPr.findall(qn('w:pBdr')):
        pPr.remove(border)

    borders = OxmlElement('w:pBdr')

    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')  # thickness
    bottom.set(qn('w:color'), hex_color.lstrip('#'))
    bottom.set(qn('w:space'), '1')

    borders.append(bottom)
    pPr.append(borders)

def add_header_with_fully_flush_left_logo(doc, logo_path, bar_color_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    section = doc.sections[0]
    header = section.header
    section.header_distance = Inches(0.15)

    for para in header.paragraphs:
        p = para._element
        p.getparent().remove(p)

    # Only add logo if logo_path is provided
    if logo_path:
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

def create_resume_doc(name, summary, certifications, skills, experience, education, logo_path=None):
    hex_teal = "#284b62"
    hex_dark = "#0b233b"

    doc = Document()
    set_strict_page_margins_fixed(doc.sections[0], inches=1)

    # Add header with logo and bar (only if logo provided)
    if logo_path:
        add_header_with_fully_flush_left_logo(doc, logo_path, hex_dark)

    # Name - Bold, larger font
    name_para = doc.add_paragraph()
    name_run = name_para.add_run(name.upper())
    name_run.font.name = 'Calibri'
    name_run.font.size = Pt(18)
    name_run.font.bold = True
    name_run.font.color.rgb = RGBColor(0, 0, 0)  # Black text
    name_para.paragraph_format.space_after = Pt(12)

    # Summary paragraph - regular text, justified
    if summary:
        summary_para = doc.add_paragraph()
        summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        summary_para.paragraph_format.space_after = Pt(18)
        summary_run = summary_para.add_run(summary)
        summary_run.font.name = 'Calibri'
        summary_run.font.size = Pt(11)

    # CERTIFICATIONS Header (separate from skills)
    if certifications:
        cert_header = doc.add_paragraph()
        cert_header_run = cert_header.add_run("CERTIFICATIONS")
        cert_header_run.font.name = 'Calibri'
        cert_header_run.font.size = Pt(14)
        cert_header_run.font.bold = True
        cert_header_run.font.color.rgb = RGBColor(0, 0, 0)
        cert_header.paragraph_format.space_after = Pt(6)

        # Certifications content - indented with > symbol like template
        cert_para = doc.add_paragraph()
        cert_para.paragraph_format.left_indent = Inches(0.5)
        cert_para.paragraph_format.space_after = Pt(18)
        cert_run = cert_para.add_run("> " + " | ".join(certifications))
        cert_run.font.name = 'Calibri'
        cert_run.font.size = Pt(11)

    # TECHNICAL SKILLS Header (separate section)
    if skills:
        skills_header = doc.add_paragraph()
        skills_header_run = skills_header.add_run("TECHNICAL SKILLS")
        skills_header_run.font.name = 'Calibri'
        skills_header_run.font.size = Pt(14)
        skills_header_run.font.bold = True
        skills_header_run.font.color.rgb = RGBColor(0, 0, 0)
        skills_header.paragraph_format.space_after = Pt(6)

        # Skills content - indented with > symbol like template
        skills_para = doc.add_paragraph()
        skills_para.paragraph_format.left_indent = Inches(0.5)
        skills_para.paragraph_format.space_after = Pt(18)
        skills_run = skills_para.add_run("> " + " | ".join(skills))
        skills_run.font.name = 'Calibri'
        skills_run.font.size = Pt(11)

    # EXPERIENCE Header with horizontal line
    exp_header = doc.add_paragraph()
    exp_header_run = exp_header.add_run("EXPERIENCE")
    exp_header_run.font.name = 'Calibri'
    exp_header_run.font.size = Pt(14)
    exp_header_run.font.bold = True
    exp_header_run.font.color.rgb = RGBColor(0, 0, 0)
    exp_header.paragraph_format.space_after = Pt(12)
    insert_horizontal_line(exp_header, "#000000")

    # Experience entries
    for job in experience:
        # Company, Location with right-aligned dates
        comp_para = doc.add_paragraph()
        comp_para.paragraph_format.space_before = Pt(6)
        comp_para.paragraph_format.space_after = Pt(0)
        
        # Company and location - bold
        comp_run = comp_para.add_run(f"{job['company']}, {job.get('city', '')}, {job.get('state', '')}")
        comp_run.font.name = 'Calibri'
        comp_run.font.size = Pt(11)
        comp_run.font.bold = True
        comp_run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Add tab and date - bold, right aligned
        comp_para.add_run("\t")
        date_run = comp_para.add_run(job['years'])
        date_run.font.name = 'Calibri'
        date_run.font.size = Pt(11)
        date_run.font.bold = True
        date_run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Set tab stop for right alignment
        comp_para.paragraph_format.tab_stops.add_tab_stop(Inches(6.0))

        # Job title - italic
        title_para = doc.add_paragraph()
        title_para.paragraph_format.space_before = Pt(0)
        title_para.paragraph_format.space_after = Pt(6)
        title_run = title_para.add_run(f"***{job['title']}***")
        title_run.font.name = 'Calibri'
        title_run.font.size = Pt(11)
        title_run.font.italic = True
        title_run.font.color.rgb = RGBColor(0, 0, 0)

        # Bullet points - using dashes like the template
        for bullet in job['bullets']:
            bullet_para = doc.add_paragraph()
            bullet_para.paragraph_format.left_indent = Inches(0.5)
            bullet_para.paragraph_format.space_before = Pt(3)
            bullet_para.paragraph_format.space_after = Pt(3)
            bullet_run = bullet_para.add_run(f"- {bullet}")
            bullet_run.font.name = 'Calibri'
            bullet_run.font.size = Pt(11)

    # EDUCATION Header
    edu_header = doc.add_paragraph()
    edu_header.paragraph_format.space_before = Pt(18)
    edu_header_run = edu_header.add_run("EDUCATION")
    edu_header_run.font.name = 'Calibri'
    edu_header_run.font.size = Pt(14)
    edu_header_run.font.bold = True
    edu_header_run.font.color.rgb = RGBColor(0, 0, 0)
    edu_header.paragraph_format.space_after = Pt(12)

    # Education entries - matching template format
    for edu in education:
        # University name - italic and bold
        univ_para = doc.add_paragraph()
        univ_para.paragraph_format.space_before = Pt(6)
        univ_run = univ_para.add_run(f"***{edu['university']}***")
        univ_run.font.name = 'Calibri'
        univ_run.font.size = Pt(11)
        univ_run.font.italic = True
        univ_run.font.bold = True
        univ_run.font.color.rgb = RGBColor(0, 0, 0)

        # Degree - italic
        deg_para = doc.add_paragraph()
        deg_para.paragraph_format.space_before = Pt(0)
        deg_para.paragraph_format.space_after = Pt(6)
        deg_run = deg_para.add_run(f"*{edu['degree']}*")
        deg_run.font.name = 'Calibri'
        deg_run.font.size = Pt(11)
        deg_run.font.italic = True
        deg_run.font.color.rgb = RGBColor(0, 0, 0)

    return doc

# Example usage matching the Jenny Oakes template format:
if __name__ == "__main__":
    # Sample data structure
    sample_data = {
        'name': 'Jenny Oakes',
        'summary': 'Dynamic and forward-thinking leader with a comprehensive background in healthcare technology, specializing in technology strategy for healthcare organizations. Proven track record in executive leadership, adeptly navigating various healthcare verticals.',
        'certifications': [
            'Project Management Professional (PMP) - Obtained 2014',
            'Multiple Epic Certifications'
        ],
        'skills': [
            'Executive presence',
            'strategic communication',
            'Cultivating and managing relationships',
            'Crafting and executing EHR strategies',
            'Revenue Cycle Management',
            'Product development',
            'innovation'
        ],
        'experience': [
            {
                'company': 'ProsperityEHR',
                'city': 'Madison',
                'state': 'WI',
                'years': '2023 to Present',
                'title': 'VP, Solutions (Revenue Cycle + Implementation Services)',
                'bullets': [
                    'Successfully launched ProsperityEHR behavioral health platform at multiple behavioral health practices.',
                    'Developed and operationalized a scalable application implementation and training framework.',
                    'Drove a 200% increase in RCM services growth, scaling from $10M to $30M AR managed in 2023.'
                ]
            }
        ],
        'education': [
            {
                'university': 'University of Minnesota, Minneapolis and Saint Paul, MN',
                'degree': 'Master of Healthcare Administration (MHA)'
            },
            {
                'university': 'University of Wisconsin, Madison, WI', 
                'degree': 'Bachelor of Arts (B.A.) in International Studies'
            }
        ]
    }
    
    # Create the document
    doc = create_resume_doc(
        sample_data['name'],
        sample_data['summary'],
        sample_data['certifications'],
        sample_data['skills'],
        sample_data['experience'],
        sample_data['education']
    )
    
    # Save the document
    doc.save("formatted_resume_fixed.docx")
