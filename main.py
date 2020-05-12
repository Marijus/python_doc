from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement
from docx.shared import Inches
from docx.shared import RGBColor
from docx.table import BlockItemContainer
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import docx
import random

document = Document()

def addCheckbox(para, box_id, name):
    run = para.add_run()
    tag = run._r
    fld = docx.oxml.shared.OxmlElement('w:fldChar')
    fld.set(docx.oxml.ns.qn('w:fldCharType'), 'begin')

    ffData = docx.oxml.shared.OxmlElement('w:ffData')
    e = docx.oxml.shared.OxmlElement('w:name')
    e.set(docx.oxml.ns.qn('w:val'), 'Check1')
    ffData.append(e)
    ffData.append(docx.oxml.shared.OxmlElement('w:enabled'))
    e = docx.oxml.shared.OxmlElement('w:calcOnExit')
    e.set(docx.oxml.ns.qn('w:val'), '0')
    ffData.append(e)
    e = docx.oxml.shared.OxmlElement('w:checkBox')
    e.append(docx.oxml.shared.OxmlElement('w:sizeAuto'))
    ee = docx.oxml.shared.OxmlElement('w:default')
    ee.set(docx.oxml.ns.qn('w:val'), '0')
    e.append(ee)
    ffData.append(e)

    fld.append(ffData)
    tag.append(fld)

    run2 = para.add_run()
    tag2 = run2._r
    start = docx.oxml.shared.OxmlElement('w:bookmarkStart')
    start.set(docx.oxml.ns.qn('w:id'), str(box_id))
    start.set(docx.oxml.ns.qn('w:name'), name)
    tag2.append(start)

    run3 = para.add_run()
    tag3 = run3._r
    instr = docx.oxml.OxmlElement('w:instrText')
    instr.text = 'FORMCHECKBOX'
    tag3.append(instr)

    run4 = para.add_run()
    tag4 = run4._r
    fld2 = docx.oxml.shared.OxmlElement('w:fldChar')
    fld2.set(docx.oxml.ns.qn('w:fldCharType'), 'end')
    tag4.append(fld2)

    run5 = para.add_run()
    tag5 = run5._r
    end = docx.oxml.shared.OxmlElement('w:bookmarkEnd')
    end.set(docx.oxml.ns.qn('w:id'), str(box_id))
    end.set(docx.oxml.ns.qn('w:name'), name)
    tag5.append(end)



def addCheckbox(para, box_id, name):

  run = para.add_run()
  tag = run._r
  fldchar = docx.oxml.shared.OxmlElement('w:fldChar')
  fldchar.set(docx.oxml.ns.qn('w:fldCharType'), 'begin')

  ffdata = docx.oxml.shared.OxmlElement('w:ffData')
  name = docx.oxml.shared.OxmlElement('w:name')
  name.set(docx.oxml.ns.qn('w:val'), cb_name)
  enabled = docx.oxml.shared.OxmlElement('w:enabled')
  calconexit = docx.oxml.shared.OxmlElement('w:calcOnExit')
  calconexit.set(docx.oxml.ns.qn('w:val'), '0')

  checkbox = docx.oxml.shared.OxmlElement('w:checkBox')
  sizeauto = docx.oxml.shared.OxmlElement('w:sizeAuto')
  default = docx.oxml.shared.OxmlElement('w:default')

  # if checked:
  #   default.set(docx.oxml.ns.qn('w:val'), '1')
  # else:
  #   default.set(docx.oxml.ns.qn('w:val'), '0')

  checkbox.append(sizeauto)
  checkbox.append(default)
  ffdata.append(name)
  ffdata.append(enabled)
  ffdata.append(calconexit)
  ffdata.append(checkbox)
  fldchar.append(ffdata)
  tag.append(fldchar)

  run2 = para.add_run()
  tag2 = run2._r
  start = docx.oxml.shared.OxmlElement('w:bookmarkStart')
  start.set(docx.oxml.ns.qn('w:id'), str(box_id))
  start.set(docx.oxml.ns.qn('w:name'), name)
  tag2.append(start)

  run3 = para.add_run()
  tag3 = run3._r
  instr = docx.oxml.OxmlElement('w:instrText')
  instr.text = 'FORMCHECKBOX'
  tag3.append(instr)

  run4 = para.add_run()
  tag4 = run4._r
  fld2 = docx.oxml.shared.OxmlElement('w:fldChar')
  fld2.set(docx.oxml.ns.qn('w:fldCharType'), 'end')
  tag4.append(fld2)

  run5 = para.add_run()
  tag5 = run5._r
  end = docx.oxml.shared.OxmlElement('w:bookmarkEnd')
  end.set(docx.oxml.ns.qn('w:id'), str(box_id))
  end.set(docx.oxml.ns.qn('w:name'), name)
  tag5.append(end)

  return

def set_cell_border(cell, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


def get_merge_cells(table,row,start,end):
    start_cell = table.cell(row, start)
    end_cell = table.cell(row, end)
    new_cell = start_cell.merge(end_cell)
    return new_cell


def get_header(header_text='RENTAL APPLICATION'):
    header = document.add_heading(header_text, 1)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_font = header.style.font
    header_font.name = 'Calibri'
    header_font.size = Pt(20)

def get_header_paragrah(text):
    header_paragraph = document.add_paragraph()
    header_paragraph_text = f'''
        Property Address: {text}
        Unit #: {text}
        City, State, ZIP: {text}
        Date of Application: {text}
    '''
    header_paragraph_font = header_paragraph.style.font
    header_paragraph_font.name = 'Calibri'
    header_paragraph_font.size = Pt(12)
    header_paragraph.add_run(header_paragraph_text).bold = True


def get_table_applicant_information():
    initials = ['First Name','','Middle Name','','Last Name','']
    contact_informations = ['Email','','Phone #1','','Phone #2','']
    document_informations = ['Date of Birth','_ _/ _ _/ _ _ _ _','Social Security #','','Driver’s License #','']
    table_applicant_information = document.add_table(rows=4,cols=6)
    header_cell = get_merge_cells(table_applicant_information,0,0,5)
    header_cell.text = 'APPLICANT INFORMATION'
    table_header_color = parse_xml(r'<w:shd {} w:fill="1F5C8B"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)

    row_second_cells = table_applicant_information.rows[1].cells
    for cell, initial in enumerate(initials):
        if cell%2 != 0:
            set_cell_border(row_second_cells[cell],bottom={"sz": 10, "color": "#00FF00", "val": "single"})
        row_second_cells[cell].text = initial
    row_third_cells = table_applicant_information.rows[2].cells
    for cell, contact_information in enumerate(contact_informations):
        if cell%2 != 0:
            set_cell_border(row_third_cells[cell],bottom={"sz": 10, "color": "#00FF00", "val": "single"})
        row_third_cells[cell].text = contact_information
    row_fourth_cells = table_applicant_information.rows[3].cells
    for cell, document_information in enumerate(document_informations):
        if cell%2 != 0:
            set_cell_border(row_fourth_cells[cell],bottom={"sz": 10, "color": "#00FF00", "val": "single"})
        row_fourth_cells[cell].text = document_information


def get_table_additional_occupant(rows=4):
    occupant_params = ['Name','','Relationship','','Age','']
    table_additional_occupant = document.add_table(rows=rows, cols=6)
    header_cell = get_merge_cells(table_additional_occupant,0,0,5)

    header_cell.text = 'ADDITIONAL OCCUPANT(S)'
    table_header_color = parse_xml(r'<w:shd {} w:fill="1F5C8B"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)

    for row in range(1,rows):
        row_cells = table_additional_occupant.rows[row].cells
        for cell, occupant_param in enumerate(occupant_params):
            if cell % 2 != 0:
                set_cell_border(row_cells[cell], bottom={"sz": 10, "color": "#00FF00", "val": "single"})
            row_cells[cell].text = occupant_param


def get_table_residence_history():
    table_residence_history = document.add_table(5,15)
    header_cell = get_merge_cells(table_residence_history, 0, 0, 14)
    header_cell.text = 'RESIDENCE HISTORY'
    table_header_color = parse_xml(r'<w:shd {} w:fill="1F5C8B"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)

    second_row = table_residence_history.rows[1].cells
    second_0_2_row = get_merge_cells(table_residence_history, 1, 0, 2)
    second_0_2_row.text = 'Current Address'
    second_3_6_row = get_merge_cells(table_residence_history, 1, 3, 6)
    set_cell_border(second_3_6_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    second_7_8_row = get_merge_cells(table_residence_history, 1, 7, 8)
    second_7_8_row.text = 'Unit #'
    second_9_10_row = get_merge_cells(table_residence_history, 1, 9, 10)
    set_cell_border(second_9_10_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    second_row[11].text = 'Rent'
    second_row[13].text = 'Own'

    third_row = table_residence_history.rows[2].cells
    third_row[0].text = 'City'
    third_1_2_row = get_merge_cells(table_residence_history, 2, 1, 2)
    set_cell_border(third_1_2_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_row[3].text = 'State'
    third_4_6_row = get_merge_cells(table_residence_history, 2, 4, 6)
    set_cell_border(third_4_6_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_row[7].text = 'ZIP'
    third_8_9_row =  get_merge_cells(table_residence_history, 2, 8, 9)
    set_cell_border(third_8_9_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_10_11_row =  get_merge_cells(table_residence_history, 2, 10, 11)
    third_10_11_row.text = 'Monthly Payment or Rent $'
    third_12_14_row = get_merge_cells(table_residence_history, 2, 12, 14)
    set_cell_border(third_12_14_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})

    fourth_0_3_row = get_merge_cells(table_residence_history, 3, 0, 3)
    fourth_0_3_row.text = 'Dates of Residence'
    fourth_4_7_row = get_merge_cells(table_residence_history, 3, 4, 7)
    fourth_4_7_row.text = '__/__/____ to __/__/____'
    set_cell_border(fourth_4_7_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    fourth_8_9_row = get_merge_cells(table_residence_history, 3, 8, 9)
    fourth_8_9_row.text ='Present Landlord'
    fourth_10_11_row = get_merge_cells(table_residence_history, 3, 10, 11)
    set_cell_border(fourth_10_11_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    fourth_11_12_row = get_merge_cells(table_residence_history, 3, 11, 12)
    fourth_11_12_row.text = 'Landlord Phone#'
    fourth_13_14_row = get_merge_cells(table_residence_history, 3, 13, 14)
    set_cell_border(fourth_13_14_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})


    fifth_0_3_row =  get_merge_cells(table_residence_history, 4, 0, 3)
    fifth_0_3_row.text = 'Reason for moving out'
    fifth_4_13_row = get_merge_cells(table_residence_history, 4, 4, 13)
    set_cell_border(fifth_4_13_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})

def get_privious_address():
    table_residence_history = document.add_table(4, 15)
    second_row = table_residence_history.rows[0].cells
    second_0_2_row = get_merge_cells(table_residence_history, 0, 0, 2)
    second_0_2_row.text = 'Previous Address'
    second_3_6_row = get_merge_cells(table_residence_history, 0, 3, 6)
    set_cell_border(second_3_6_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    second_7_8_row = get_merge_cells(table_residence_history, 0, 7, 8)
    second_7_8_row.text = 'Unit #'
    second_9_10_row = get_merge_cells(table_residence_history, 0, 9, 10)
    set_cell_border(second_9_10_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    second_row[11].text = 'Rent'
    second_row[13].text = 'Own'

    third_row = table_residence_history.rows[1].cells
    third_row[0].text = 'City'
    third_1_2_row = get_merge_cells(table_residence_history, 1, 1, 2)
    set_cell_border(third_1_2_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_row[3].text = 'State'
    third_4_6_row = get_merge_cells(table_residence_history, 1, 4, 6)
    set_cell_border(third_4_6_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_row[7].text = 'ZIP'
    third_8_9_row = get_merge_cells(table_residence_history, 1, 8, 9)
    set_cell_border(third_8_9_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    third_10_11_row = get_merge_cells(table_residence_history, 1, 10, 11)
    third_10_11_row.text = 'Monthly Payment or Rent $'
    third_12_14_row = get_merge_cells(table_residence_history, 1, 12, 14)
    set_cell_border(third_12_14_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})

    fourth_0_3_row = get_merge_cells(table_residence_history, 2, 0, 3)
    fourth_0_3_row.text = 'Dates of Residence'
    fourth_4_7_row = get_merge_cells(table_residence_history, 2, 4, 7)
    fourth_4_7_row.text = '__/__/____ to __/__/____'
    set_cell_border(fourth_4_7_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    fourth_8_9_row = get_merge_cells(table_residence_history, 2, 8, 9)
    fourth_8_9_row.text = 'Previous Landlord'
    fourth_10_11_row = get_merge_cells(table_residence_history, 2, 10, 11)
    set_cell_border(fourth_10_11_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})
    fourth_11_12_row = get_merge_cells(table_residence_history, 2, 11, 12)
    fourth_11_12_row.text = 'Landlord Phone#'
    fourth_13_14_row = get_merge_cells(table_residence_history, 2, 13, 14)
    set_cell_border(fourth_13_14_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})

    fifth_0_3_row = get_merge_cells(table_residence_history, 3, 0, 3)
    fifth_0_3_row.text = 'Reason for moving out'
    fifth_4_13_row = get_merge_cells(table_residence_history, 3, 4, 13)
    set_cell_border(fifth_4_13_row, bottom={"sz": 10, "color": "#00FF00", "val": "single"})


get_header()
get_header_paragrah('oleh')
get_table_applicant_information()
get_table_additional_occupant()
get_table_residence_history()
get_privious_address()
document.save('test1.docx')

