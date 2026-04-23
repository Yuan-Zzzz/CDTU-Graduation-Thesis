import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

filename = sys.argv[1]
doc = Document(filename)

STYLE_MAPPING = {
    'Heading 1': '73',
    'Heading 2': '74',
    'Heading 3': '75',
    'FirstParagraph': '76',
    'BodyText': '76',
    'Compact': '76',
    'TOCHeading': '68',
    'toc 1': '68',
    'toc 2': '69',
    'toc 3': '70',
    'Bibliography': '76',
}

def get_style_id_by_name(doc, style_name):
    for style in doc.styles:
        if style.name == style_name:
            return style.style_id
    return None

def apply_style_by_id(paragraph, style_id):
    pPr = paragraph._element.get_or_add_pPr()
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        pStyle = OxmlElement('w:pStyle')
        if len(pPr) > 0:
            pPr.insert(0, pStyle)
        else:
            pPr.append(pStyle)
    pStyle.set(qn('w:val'), style_id)

for p in doc.paragraphs:
    pPr = p._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    raw_style_id = None
    if pPr is not None:
        pStyle = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        if pStyle is not None:
            raw_style_id = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    
    current_style_name = p.style.name if p.style else 'Normal'
    
    if raw_style_id in STYLE_MAPPING:
        target_style_id = STYLE_MAPPING[raw_style_id]
        apply_style_by_id(p, target_style_id)
        print(f"映射: raw='{raw_style_id}' -> styleId={target_style_id}")
    elif current_style_name in STYLE_MAPPING:
        target_style_id = STYLE_MAPPING[current_style_name]
        apply_style_by_id(p, target_style_id)
        print(f"映射: name='{current_style_name}' -> styleId={target_style_id}")
    else:
        if current_style_name and not current_style_name.startswith('Normal'):
            template_style_id = get_style_id_by_name(doc, current_style_name)
            if template_style_id:
                apply_style_by_id(p, template_style_id)
                print(f"保留: '{current_style_name}' -> styleId={template_style_id}")

doc.save(filename)
print(f"样式修复完成: {filename}")
