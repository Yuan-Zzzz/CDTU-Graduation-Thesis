import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

filename = sys.argv[1]
doc = Document(filename)

for p in doc.paragraphs:
    if p.style.name == 'Heading 1':
        # Change style to Normal to remove it from TOC
        p.style = doc.styles['Normal']
        
        # Apply direct formatting to match heading 1
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add spacing before and after (480 twips = 24 pt before, 240 twips = 12 pt after)
        pPr = p._element.get_or_add_pPr()
        spacing = pPr.get_or_add_spacing()
        spacing.set(qn('w:before'), '480')
        spacing.set(qn('w:after'), '240')
        
        # Format the text
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(16) # 32 half-points = 16 pt (三号)
            
            # Set font to HeiTi
            r = run._element
            rPr = r.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            rFonts.set(qn('w:eastAsia'), '黑体')
            rFonts.set(qn('w:cs'), '黑体')

doc.save(filename)
