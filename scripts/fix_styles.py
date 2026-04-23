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

ABSTRACT_TITLE_STYLES = {
    '摘要': '61',
    'Abstract': '64',
}

ABSTRACT_CONTENT_STYLES = {
    '中文摘要 内容': '62',
    '英文摘要 内容': '65',
}

ABSTRACT_KEYWORD_STYLES = {
    '关键词': '63',
    'Keywords': '66',
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

def is_chinese_text(text):
    for char in text:
        if '\u4e00' <= char <= '\u9fff':
            return True
    return False

def detect_abstract_context(paragraphs, current_idx):
    if current_idx < 0 or current_idx >= len(paragraphs):
        return None
    
    p = paragraphs[current_idx]
    text = p.text.strip()
    
    if text in ABSTRACT_TITLE_STYLES:
        return 'title', ABSTRACT_TITLE_STYLES[text]
    
    if text.startswith('关键词：') or text.startswith('Keywords:'):
        return 'keywords', ABSTRACT_KEYWORD_STYLES.get('关键词' if text.startswith('关键词') else 'Keywords')

    return None


def split_keyword_paragraph(paragraph):
    text = paragraph.text
    if text.startswith('关键词：'):
        label = '关键词：'
        content = text[len(label):]
    elif text.startswith('Keywords:'):
        label = 'Keywords:'
        content = text[len(label):]
    else:
        return False

    p = paragraph._element
    pPr = p.find(qn('w:pPr'))

    p.clear()
    if pPr is not None:
        p.append(pPr)

    r1 = OxmlElement('w:r')
    rPr1 = OxmlElement('w:rPr')
    b1 = OxmlElement('w:b')
    b1Cs = OxmlElement('w:bCs')
    rPr1.append(b1)
    rPr1.append(b1Cs)
    r1.append(rPr1)
    t1 = OxmlElement('w:t')
    t1.text = label
    r1.append(t1)
    p.append(r1)

    r2 = OxmlElement('w:r')
    t2 = OxmlElement('w:t')
    t2.text = content
    r2.append(t2)
    p.append(r2)

    return True

def make_run_bold(paragraph):
    for run in paragraph.runs:
        r = run._element
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            r.insert(0, rPr)
        if rPr.find(qn('w:b')) is None:
            rPr.append(OxmlElement('w:b'))
        if rPr.find(qn('w:bCs')) is None:
            rPr.append(OxmlElement('w:bCs'))

for i, p in enumerate(doc.paragraphs):
    pPr = p._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    raw_style_id = None
    if pPr is not None:
        pStyle = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        if pStyle is not None:
            raw_style_id = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
    
    current_style_name = p.style.name if p.style else 'Normal'
    text = p.text.strip()
    
    abstract_context = detect_abstract_context(doc.paragraphs, i)
    
    if abstract_context:
        context_type, target_style_id = abstract_context
        apply_style_by_id(p, target_style_id)
        if context_type == 'keywords':
            split_keyword_paragraph(p)
        elif context_type == 'title' and text == 'Abstract':
            make_run_bold(p)
        print(f"摘要映射: {context_type}='{text[:20]}' -> styleId={target_style_id}")
    elif raw_style_id in STYLE_MAPPING:
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
