from __future__ import print_function
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from docx.oxml import ns
import cx_Oracle,platform,docx,psutil,os,socket
#print (os.uname())
import matplotlib.pyplot as plt
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.enum.dml import MSO_THEME_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
##Define page number
def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)

def shade_cells(cells, shade):
	for cell in cells:
		tcPr = cell._tc.get_or_add_tcPr()
		tcVAlign = OxmlElement("w:shd")
		tcVAlign.set(qn("w:fill"), shade)
		tcPr.append(tcVAlign)

def add_page_number(run):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

doc = docx.Document('/app/daac/orcl_lld_daac.docx')
style = doc.styles['Body Text']
font = style.font
font.name = 'Verdana'
font.size = docx.shared.Pt(9)
today = date.today()
v_dt = today.strftime("%m/%d/%Y")
v_dt1 = today.strftime("%Y%m%d")
