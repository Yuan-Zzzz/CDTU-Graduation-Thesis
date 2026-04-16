import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def fix_citations(filename):
    doc = Document(filename)
    body = doc._element.body

    # 1. Move bookmarks inside paragraphs
    starts = body.xpath('.//w:bookmarkStart')
    ends = body.xpath('.//w:bookmarkEnd')

    for start in starts:
        name = start.get(qn('w:name'))
        if name and name.startswith('ref-'):
            p = start.getnext()
            if p is not None and p.tag == qn('w:p'):
                p.insert(0, start)

    for end in ends:
        p = end.getprevious()
        if p is not None and p.tag == qn('w:p'):
            p.append(end)

    # 2. Fix hyperlinks
    for hl in body.xpath('.//w:hyperlink'):
        anchor = hl.get(qn('w:anchor'))
        if anchor and anchor.startswith('ref-'):
            hl.set(qn('w:history'), '1')
            
            # Check previous sibling
            prev_r = hl.getprevious()
            if prev_r is not None and prev_r.tag == qn('w:r'):
                t_els = prev_r.xpath('.//w:t')
                if len(t_els) == 1 and t_els[0].text == '[':
                    # Add Hyperlink style
                    rPr = prev_r.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = OxmlElement('w:rPr')
                        prev_r.insert(0, rPr)
                    rStyle = rPr.find(qn('w:rStyle'))
                    if rStyle is None:
                        rStyle = OxmlElement('w:rStyle')
                        rPr.append(rStyle)
                    rStyle.set(qn('w:val'), '22') # Hyperlink style
                    
                    hl.insert(0, prev_r)
            
            # Check next sibling
            next_r = hl.getnext()
            if next_r is not None and next_r.tag == qn('w:r'):
                t_els = next_r.xpath('.//w:t')
                if len(t_els) == 1 and t_els[0].text == ']':
                    # Add Hyperlink style
                    rPr = next_r.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = OxmlElement('w:rPr')
                        next_r.insert(0, rPr)
                    rStyle = rPr.find(qn('w:rStyle'))
                    if rStyle is None:
                        rStyle = OxmlElement('w:rStyle')
                        rPr.append(rStyle)
                    rStyle.set(qn('w:val'), '22') # Hyperlink style
                    
                    hl.append(next_r)

    doc.save(filename)

if __name__ == '__main__':
    fix_citations(sys.argv[1])
