from xlrd import open_workbook,XL_CELL_TEXT,XLRDError
import xlsxwriter
from xlrd import open_workbook
import os,sys
import shutil
from docx import Document
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from xlrd import open_workbook,XL_CELL_TEXT,XLRDError


def MarkIndexEntry(entry,paragraph):
    run = paragraph.add_run()
    r = run._r
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar)

    run = paragraph.add_run()
    r = run._r
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' XE "%s" '%(entry)
    r.append(instrText)

    run = paragraph.add_run()
    r = run._r
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    r.append(fldChar)



def Figure(paragraph):
    paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = run = paragraph.add_run()
    r = run._r
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar)
    instrText = OxmlElement('w:instrText')
    instrText.text = ' SEQ Figure \* ARABIC'
    r.append(instrText)
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    r.append(fldChar)


def Table(paragraph):
    run = run = paragraph.add_run()
    r = run._r
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar)
    instrText = OxmlElement('w:instrText')
    instrText.text = ' SEQ Table \* ARABIC'
    r.append(instrText)
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    r.append(fldChar)

#
# The main program
#
document = Document()
paragraph = document.add_paragraph('')

paragraph = document.add_paragraph('Figure ', style='Caption')
Figure(paragraph)
paragraph.add_run(' Figure Caption ')
figpa = document.add_paragraph()
figpa.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
run = figpa.add_run()
run.add_picture('image6.png', width=Cm(10))

paragraph = document.add_paragraph('Figure Text', style='BodyText')
document.add_page_break()

document.save(str.strip('Test-Figures.docx'))
