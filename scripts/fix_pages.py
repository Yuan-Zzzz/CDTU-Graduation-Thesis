from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
FINAL_DOC = REPO_ROOT / "output" / "Final_thesis.docx"

doc = Document(str(FINAL_DOC))
body = doc._element.body

# 1. Find the sdt element (TOC) and insert a page break before it
sdt_element = None
for child in body:
    if child.tag.endswith('}sdt'):
        sdt_element = child
        break

if sdt_element is not None:
    # Create a page break paragraph
    pb_p = OxmlElement('w:p')
    pb_r = OxmlElement('w:r')
    pb_br = OxmlElement('w:br')
    pb_br.set(qn('w:type'), 'page')
    pb_r.append(pb_br)
    pb_p.append(pb_r)
    sdt_element.addprevious(pb_p)

# 2. Find the first paragraph of Chapter 1
chap1_p = None
for p in doc.paragraphs:
    if p.text.startswith('第一章'):
        chap1_p = p._element
        break

if chap1_p is not None:
    # Create a new paragraph for the section break
    sect_p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    sect_p.append(pPr)
    
    # Copy the document's main sectPr
    doc_sectPr = body.sectPr
    new_sectPr = copy.deepcopy(doc_sectPr)
    
    # Modify it to be a next page section break
    type_el = new_sectPr.find(qn('w:type'))
    if type_el is None:
        type_el = OxmlElement('w:type')
        new_sectPr.append(type_el)
    type_el.set(qn('w:val'), 'nextPage')
    
    # Set Roman numerals for the front matter (this section)
    pgNumType_front = new_sectPr.find(qn('w:pgNumType'))
    if pgNumType_front is None:
        pgNumType_front = OxmlElement('w:pgNumType')
        new_sectPr.append(pgNumType_front)
    pgNumType_front.set(qn('w:fmt'), 'upperRoman')
    pgNumType_front.set(qn('w:start'), '1')
    
    pPr.append(new_sectPr)
    
    # Insert before Chapter 1
    chap1_p.addprevious(sect_p)
    
    # Update the document's main sectPr (main matter) to restart page numbering at 1 with decimal format
    pgNumType_main = doc_sectPr.find(qn('w:pgNumType'))
    if pgNumType_main is None:
        pgNumType_main = OxmlElement('w:pgNumType')
        doc_sectPr.append(pgNumType_main)
    pgNumType_main.set(qn('w:fmt'), 'decimal')
    pgNumType_main.set(qn('w:start'), '1')

doc.save(str(FINAL_DOC))
