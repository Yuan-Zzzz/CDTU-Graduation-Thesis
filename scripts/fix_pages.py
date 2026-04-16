from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.parts.hdrftr import FooterPart
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

# 3. Modify the footer for the main body (Section 1 onwards) to display "第xx页"
if len(doc.sections) > 1:
    # We only need to create the new footer part once, then we can link it to all subsequent sections
    orig_footer_part = doc.sections[1].footer.part
    
    # Create a new footer XML by parsing the original one
    new_footer_xml = copy.deepcopy(orig_footer_part._element)
    
    # Modify the new footer XML to add "第 " and " 页"
    begins = new_footer_xml.xpath('.//w:fldChar[@w:fldCharType="begin"]')
    ends = new_footer_xml.xpath('.//w:fldChar[@w:fldCharType="end"]')
    
    if begins and ends:
        r_prefix = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>第 </w:t></w:r>')
        r_suffix = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t> 页</w:t></w:r>')
        
        begin_r = begins[0].getparent()
        begin_r.addprevious(r_prefix)
        
        end_r = ends[0].getparent()
        end_r.addnext(r_suffix)
    
    # Create a new FooterPart
    new_partname = doc.part.package.next_partname(PackURI('/word/footer%d.xml'))
    new_footer_part = FooterPart(new_partname, orig_footer_part.content_type, new_footer_xml, doc.part.package)
    
    # Relate the document part to the new footer part
    rel_id = doc.part.relate_to(new_footer_part, RT.FOOTER)
    
    # Update the sectPr for all sections starting from Section 1 to point to the new footer
    for i in range(1, len(doc.sections)):
        sectPr = doc.sections[i]._sectPr
        # Find the existing footerReference
        footer_refs = sectPr.findall(qn('w:footerReference'))
        for ref in footer_refs:
            # We only replace the default footer
            if ref.get(qn('w:type')) == 'default':
                ref.set(qn('r:id'), rel_id)

doc.save(str(FINAL_DOC))
