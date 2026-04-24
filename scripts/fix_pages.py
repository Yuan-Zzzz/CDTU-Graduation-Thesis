from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.parts.hdrftr import FooterPart, HeaderPart
import copy
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
FINAL_DOC = REPO_ROOT / "output" / "Final_thesis.docx"

doc = Document(str(FINAL_DOC))
body = doc._element.body

def insert_page_break_before(element):
    pb_p = OxmlElement('w:p')
    pb_r = OxmlElement('w:r')
    pb_br = OxmlElement('w:br')
    pb_br.set(qn('w:type'), 'page')
    pb_r.append(pb_br)
    pb_p.append(pb_r)
    element.addprevious(pb_p)

sdt_element = None
for child in body:
    if child.tag.endswith('}sdt'):
        sdt_element = child
        break

if sdt_element is not None:
    insert_page_break_before(sdt_element)

for p in doc.paragraphs:
    text = p.text.strip()
    style = p.style.name if p.style else ''
    
    if text == '摘要' and '中文摘要' in style:
        insert_page_break_before(p._element)
    elif text == 'Abstract' and '英文摘要' in style:
        insert_page_break_before(p._element)

for p in doc.paragraphs:
    text = p.text.strip()
    if text.startswith('第') and '章' in text and len(text) < 30:
        insert_page_break_before(p._element)
    elif text == '参考文献':
        insert_page_break_before(p._element)
    elif text == '致谢' or (text.startswith('致') and '谢' in text and len(text) < 10):
        insert_page_break_before(p._element)

chap1_p = None
for p in doc.paragraphs:
    if p.text.startswith('第一章') or p.text.startswith('第 1 章'):
        chap1_p = p._element
        break

if chap1_p is not None:
    sect_p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    sect_p.append(pPr)
    
    doc_sectPr = body.sectPr
    new_sectPr = copy.deepcopy(doc_sectPr)
    
    type_el = new_sectPr.find(qn('w:type'))
    if type_el is None:
        type_el = OxmlElement('w:type')
        new_sectPr.append(type_el)
    type_el.set(qn('w:val'), 'nextPage')
    
    pgNumType_front = new_sectPr.find(qn('w:pgNumType'))
    if pgNumType_front is None:
        pgNumType_front = OxmlElement('w:pgNumType')
        new_sectPr.append(pgNumType_front)
    pgNumType_front.set(qn('w:fmt'), 'upperRoman')
    pgNumType_front.set(qn('w:start'), '1')
    
    pPr.append(new_sectPr)
    chap1_p.addprevious(sect_p)
    
    pgNumType_main = doc_sectPr.find(qn('w:pgNumType'))
    if pgNumType_main is None:
        pgNumType_main = OxmlElement('w:pgNumType')
        doc_sectPr.append(pgNumType_main)
    pgNumType_main.set(qn('w:fmt'), 'decimal')
    pgNumType_main.set(qn('w:start'), '1')

header_xml_str = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14 w15 wp14">
  <w:p w14:paraId="46D6B860">
    <w:pPr>
      <w:pStyle w:val="77"/>
      <w:rPr>
        <w:color w:val="FF0000"/>
      </w:rPr>
    </w:pPr>
    <w:r>
      <w:t>成都工业学院本科毕业论文（设计）</w:t>
    </w:r>
  </w:p>
</w:hdr>'''

header_element = parse_xml(header_xml_str.encode('utf-8'))
new_partname = doc.part.package.next_partname(PackURI('/word/header%d.xml'))
new_header_part = HeaderPart(new_partname, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', header_element, doc.part.package)
header_rel_id = doc.part.relate_to(new_header_part, RT.HEADER)

if len(doc.sections) > 1:
    orig_footer_part = doc.sections[1].footer.part
    new_footer_xml = copy.deepcopy(orig_footer_part._element)
    
    begins = new_footer_xml.xpath('.//w:fldChar[@w:fldCharType="begin"]')
    ends = new_footer_xml.xpath('.//w:fldChar[@w:fldCharType="end"]')
    
    if begins and ends:
        r_prefix = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>第 </w:t></w:r>')
        r_suffix = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t> 页</w:t></w:r>')
        
        begin_r = begins[0].getparent()
        begin_r.addprevious(r_prefix)
        
        end_r = ends[0].getparent()
        end_r.addnext(r_suffix)
    
    footer_partname = doc.part.package.next_partname(PackURI('/word/footer%d.xml'))
    new_footer_part = FooterPart(footer_partname, orig_footer_part.content_type, new_footer_xml, doc.part.package)
    footer_rel_id = doc.part.relate_to(new_footer_part, RT.FOOTER)
    
    for i in range(1, len(doc.sections)):
        sectPr = doc.sections[i]._sectPr
        
        header_ref = sectPr.find(qn('w:headerReference'))
        if header_ref is None:
            header_ref = OxmlElement('w:headerReference')
            sectPr.insert(0, header_ref)
        header_ref.set(qn('w:type'), 'default')
        header_ref.set(qn('r:id'), header_rel_id)
        
        footer_refs = sectPr.findall(qn('w:footerReference'))
        for ref in footer_refs:
            if ref.get(qn('w:type')) == 'default':
                ref.set(qn('r:id'), footer_rel_id)

doc.save(str(FINAL_DOC))
