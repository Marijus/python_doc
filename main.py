from textwrap import dedent

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT
from docx.shared import Pt, Cm
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement
from docx.shared import Inches
from docx.shared import RGBColor
from docx.table import BlockItemContainer, _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import docx
import random



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

def set_cell_margins(cell: _Cell, **kwargs):
    """
    cell:  actual cell instance you want to modify

    usage:

        set_cell_margins(cell, top=50, start=50, bottom=50, end=50)

    provided values are in twentieths of a point (1/1440 of an inch).
    read more here: http://officeopenxml.com/WPtableCellMargins.php
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for m in [
        "top",
        "start",
        "bottom",
        "end",
    ]:
        if m in kwargs:
            node = OxmlElement("w:{}".format(m))
            node.set(qn('w:w'), str(kwargs.get(m)))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)

    tcPr.append(tcMar)

def get_merge_cells(table,row,start,end):
    start_cell = table.cell(row, start)
    end_cell = table.cell(row, end)
    new_cell = start_cell.merge(end_cell)
    return new_cell


def get_header(header_text='RENTAL APPLICATION'):
    header = document.add_heading(header_text, 1)
    header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    header_font = header.style.font
    header_font.name = 'Calibri'
    header_font.size = Pt(20)


def get_header_paragraph(text):
    header_paragraph = document.add_paragraph()
    header_paragraph_text = f'''
    Property Address: {text}
    Unit #: {text}
    City, State, ZIP: {text}
    Date of Application: {text}
    '''
    # header_paragraph.aligmnet = WD_TAB_ALIGNMENT
    header_paragraph_font = header_paragraph.style.font
    header_paragraph_font.name = 'Calibri'
    header_paragraph_font.size = Pt(12)
    header_paragraph.add_run(dedent(header_paragraph_text)).bold = True


def get_table_applicant_information():

    initials = ['First Name','','Middle Name','','Last Name','']
    contact_informations = ['Email','','Phone #1','','Phone #2','']
    document_informations = ['Date of Birth','_ _/ _ _/ _ _ _ _','Social Security #','','Driver’s License #','']

    table_applicant_information = document.add_table(rows=4,cols=6)

    # table_applicant_information.autofit = False
    # for row in table_applicant_information.rows:
    #     row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    #     row.height = Cm(0.6)

    header_cell = get_merge_cells(table_applicant_information,0,0,5)
    header_cell.text = 'APPLICANT INFORMATION'
    table_header_color = parse_xml(r'<w:shd {} w:fill="#1f3864"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)

    for row in range(4):
        for cell in range(6):
            table_applicant_information.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
            if cell % 2 != 0:
                set_cell_border(table_applicant_information.cell(row, cell), bottom={"sz": 6, "color": "#000000", "val": "single"})

    row_second_cells = table_applicant_information.rows[1].cells
    for cell, initial in enumerate(initials):
        row_second_cells[cell].text = initial
        row_second_cells[cell].paragraphs[0].paragraph_format.space_after = Pt(0)

    row_third_cells = table_applicant_information.rows[2].cells
    for cell, contact_information in enumerate(contact_informations):
        row_third_cells[cell].text = contact_information
        # row_third_cells[cell].paragraphs[0].paragraph_format.space_after = Pt(0)

    row_fourth_cells = table_applicant_information.rows[3].cells
    for cell, document_information in enumerate(document_informations):
        row_fourth_cells[cell].text = document_information
        row_fourth_cells[cell].paragraphs[0].paragraph_format.space_after = Pt(0)

    header_paragraph = document.add_paragraph()
    header_paragraph.aligmnet = WD_PARAGRAPH_ALIGNMENT.CENTER
    header_paragraph_font = header_paragraph.style.font
    header_paragraph_font.name = 'Calibri'
    header_paragraph_font.size = Pt(12)


def get_table_additional_occupant(rows=4):
    occupant_params = ['Name','','Relationship','','Age','']
    table_additional_occupant = document.add_table(rows=rows, cols=7)

    table_additional_occupant.autofit = False

    for row in table_additional_occupant.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(0.60)



    paragraph = document.add_paragraph()
    paragraph.aligmnet = WD_TAB_ALIGNMENT
    paragraph_font = paragraph.style.font
    paragraph_font.name = 'Calibri'
    paragraph_font.size = Pt(10)


    header_cell = get_merge_cells(table_additional_occupant,0,0,6)
    header_cell.text = 'ADDITIONAL OCCUPANT(S)'
    table_header_color = parse_xml(r'<w:shd {} w:fill="#1f3864"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)


    for row in range(1, rows):
        column = table_additional_occupant.rows[row].cells
        column[0].text = 'Name'
        column[0].width = Inches(0.5)
        column_1_2 = get_merge_cells(table_additional_occupant, row, 1, 2)
        set_cell_border(column_1_2, bottom={"sz": 6, "color": "#000000", "val": "single"})

        column[3].text = 'Relationship'
        column[3].width = Inches(0.9)
        set_cell_border(column[4], bottom={"sz": 6, "color": "#000000", "val": "single"})

        column[5].text = 'Age'
        column[5].width = Inches(0.5)
        set_cell_border(column[6], bottom={"sz": 6, "color": "#000000", "val": "single"})

#1f3864
def get_table_residence_history():
    table_residence_history = document.add_table(5,15)
    header_cell = get_merge_cells(table_residence_history, 0, 0, 14)
    header_cell.text = 'RESIDENCE HISTORY'
    table_header_color = parse_xml(r'<w:shd {} w:fill="#1f3864"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)

    table_residence_history.autofit = False
    for row in table_residence_history.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(0.64)

    for row in range(1,5):
        row_cell = table_residence_history.rows[row].cells
        for cell in range(0,2):
            row_cell[cell].width = Inches(0.4)

    for row in range(5):
        for cell in range(15):
            table_residence_history.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM

    paragraph = document.add_paragraph()
    paragraph.aligmnet = WD_TAB_ALIGNMENT
    paragraph_font = paragraph.style.font
    paragraph_font.name = 'Calibri'
    paragraph_font.size = Pt(10)


    second_row = table_residence_history.rows[1].cells
    second_0_2_row = get_merge_cells(table_residence_history, 1, 0, 2)
    print(second_0_2_row.width.inches)
    second_0_2_row.text = 'Current Address'
    second_0_2_row.width = Inches(1.2)
    second_3_7_row = get_merge_cells(table_residence_history, 1, 3, 7)
    set_cell_border(second_3_7_row, bottom={"sz": 6, "color": "#000000", "val": "single"},)
    second_row[8].text = 'Unit #'
    second_9_10_row = get_merge_cells(table_residence_history, 1, 9, 10)
    set_cell_border(second_9_10_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    second_row[11].text = 'Rent'
    second_row[12].text = u"  □  "
    second_row[13].text = 'Own'
    second_row[14].text = u"  □  "


    third_row = table_residence_history.rows[2].cells
    third_row[0].text = 'City'
    third_1_2_row = get_merge_cells(table_residence_history, 2, 1, 2)
    set_cell_border(third_1_2_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_row[3].text = 'State'
    third_4_5_row = get_merge_cells(table_residence_history, 2, 4, 5)
    set_cell_border(third_4_5_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_row[6].text = 'ZIP'
    third_7_8_row = get_merge_cells(table_residence_history, 2, 7, 8)
    set_cell_border(third_7_8_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_9_12_row =  get_merge_cells(table_residence_history, 2, 9, 12)
    third_9_12_row.text = 'Monthly Payment or Rent $'
    third_13_14_row = get_merge_cells(table_residence_history, 2, 13, 14)
    set_cell_border(third_13_14_row, bottom={"sz": 6, "color": "#000000", "val": "single"})



    fourth_row = table_residence_history.rows[3].cells
    fourth_0_2_row = get_merge_cells(table_residence_history, 3, 0, 2)
    fourth_0_2_row.text = 'Dates of Residence'
    fourth_row[3].width = Inches(0.75)
    fourth_row[4].width = Inches(0.65)
    fourth_3_5_row = get_merge_cells(table_residence_history, 3, 3, 5)
    print(fourth_3_5_row.width.inches)
    fourth_3_5_row.text = f'__/__/____ to __/__/____'
    set_cell_border(fourth_3_5_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    fourth_row[6].width = Inches(0.1)
    fourth_6_8_row = get_merge_cells(table_residence_history, 3, 6, 8)
    fourth_6_8_row.text ='Present Landlord'
    set_cell_border(fourth_row[9], bottom={"sz": 6, "color": "#000000", "val": "single"})
    fourth_row[9].width = Inches(0.5)
    fourth_10_12_row = get_merge_cells(table_residence_history, 3, 10, 12)
    fourth_10_12_row.text = 'Landlord Phone#'
    fourth_row[13].width = Inches(0.25)
    fourth_row[14].width = Inches(0.25)
    fourth_13_14_row = get_merge_cells(table_residence_history, 3, 13, 14)

    set_cell_border(fourth_13_14_row, bottom={"sz": 6, "color": "#000000", "val": "single"})


    fifth_0_2_row =  get_merge_cells(table_residence_history, 4, 0, 2)
    fifth_0_2_row.text = 'Reason for moving out'
    fifth_3_14_row = get_merge_cells(table_residence_history, 4, 3, 14)
    set_cell_border(fifth_3_14_row, bottom={"sz": 6, "color": "#000000", "val": "single"})

def get_privious_address():
    table_residence_history = document.add_table(4, 15)
    second_row = table_residence_history.rows[0].cells
    second_0_2_row = get_merge_cells(table_residence_history, 0, 0, 2)
    second_0_2_row.text = 'Previous Address'
    second_3_6_row = get_merge_cells(table_residence_history, 0, 3, 6)
    set_cell_border(second_3_6_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    second_7_8_row = get_merge_cells(table_residence_history, 0, 7, 8)
    second_7_8_row.text = 'Unit #'
    second_9_10_row = get_merge_cells(table_residence_history, 0, 9, 10)
    set_cell_border(second_9_10_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    second_row[11].text = 'Rent'
    second_row[13].text = 'Own'

    third_row = table_residence_history.rows[1].cells
    third_row[0].text = 'City'
    third_1_2_row = get_merge_cells(table_residence_history, 1, 1, 2)
    set_cell_border(third_1_2_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_row[3].text = 'State'
    third_4_6_row = get_merge_cells(table_residence_history, 1, 4, 6)
    set_cell_border(third_4_6_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_row[7].text = 'ZIP'
    third_8_9_row = get_merge_cells(table_residence_history, 1, 8, 9)
    set_cell_border(third_8_9_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    third_10_11_row = get_merge_cells(table_residence_history, 1, 10, 11)
    third_10_11_row.text = 'Monthly Payment or Rent $'
    third_12_14_row = get_merge_cells(table_residence_history, 1, 12, 14)
    set_cell_border(third_12_14_row, bottom={"sz": 6, "color": "#000000", "val": "single"})


    fourth_0_3_row = get_merge_cells(table_residence_history, 2, 0, 3)
    fourth_0_3_row.text = 'Dates of Residence'
    fourth_4_6_row = get_merge_cells(table_residence_history, 2, 4, 6)
    fourth_4_6_row.text = '__/__/____ to __/__/____'
    set_cell_border(fourth_4_6_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    fourth_7_8_row = get_merge_cells(table_residence_history, 2, 7, 8)
    fourth_7_8_row.text = 'Previous Landlord'
    fourth_9_10_row = get_merge_cells(table_residence_history, 2, 9, 10)
    set_cell_border(fourth_9_10_row, bottom={"sz": 6, "color": "#000000", "val": "single"})
    fourth_11_13_row = get_merge_cells(table_residence_history, 2, 11, 13)
    fourth_11_13_row.text = 'Landlord Phone#'
    fourth_14_row = get_merge_cells(table_residence_history, 2, 13, 14)
    set_cell_border(fourth_14_row, bottom={"sz": 6, "color": "#000000", "val": "single"})

    fifth_0_3_row = get_merge_cells(table_residence_history, 3, 0, 3)
    fifth_0_3_row.text = 'Reason for moving out'
    fifth_4_13_row = get_merge_cells(table_residence_history, 3, 4, 13)
    set_cell_border(fifth_4_13_row, bottom={"sz": 6, "color": "#000000", "val": "single"})

def get_additionals_questions_text():
    header_paragraph = document.add_paragraph()

    text = '''
    It is against a law to discriminate against any person in the terms, conditions, or privileges of sale or rental of 
    a dwelling, or in the provision of services or facilities in connection therewith, because of race, color, religion,
     sex, familial status, or national origin.
    
    By filling this form, applicant authorizes the verification of all information provided in this application,
     including Social Security Number, employment and income, credit history, previous and current rental history and 
     any other relevant information necessary for the Landlord to evaluate the application. 
     
    Applicant agrees that false or incomplete information filled in this application may result in a rejection 
    of this application and/or termination of a rental agreement.
    Non-refundable application fee: $20.00 
    '''
    header_paragraph.aligmnet = WD_TAB_ALIGNMENT
    header_paragraph_font = header_paragraph.style.font
    header_paragraph_font.name = 'Calibri'
    header_paragraph_font.size = Pt(12)
    header_paragraph.add_run(dedent(text)).bold = True

def get_table_sign():
    table_sign = document.add_table(4,1)
    table_sign.autofit = False
    for number_row,row in enumerate(table_sign.rows):
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.width = Inches(4)
        if number_row == 1 or number_row == 4 :
            row.height = Cm(0.64)
        else:
            row.height = Cm(1.8)
    for row in range(4):
        table_sign.rows[row].cells[0].width = Inches(2.5)
    table_sign.rows[0].cells[0].text = ''
    signature = table_sign.rows[1].cells[0].add_paragraph('Landlord Signature')
    signature.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # table_sign.rows[1].cells[0].text = 'Landlord Signature'
    table_sign.rows[2].cells[0].text = ''
    # set_cell_border(table_sign.rows[2].cells[0], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    table_sign.rows[3].cells[0].text = 'Date'


def get_table_employment_information():

    table_employment_information = document.add_table(rows=5, cols=8)

    header_cell = get_merge_cells(table_employment_information, 0, 0, 7)
    header_cell.text = 'employment information'.upper()
    table_header_color = parse_xml(r'<w:shd {} w:fill="#1f3864"/>'.format(nsdecls('w')))
    header_cell._tc.get_or_add_tcPr().append(table_header_color)
    header_cell.paragraphs[0].paragraph_format.space_after = Pt(0)

    employment_information = ['Current Employer', '', 'Position/Title', '']

    second_row = table_employment_information.rows[1]
    second_row.cells[0].text = 'Current Employer'
    second_row.cells[0].width = Inches(1.4)
    second_row.cells[1].width = Inches(2.6)
    set_cell_border(second_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 1, 2, 3)
    second_row.cells[2].text = 'Position/Title'
    second_row.cells[2].width = Inches(1.4)
    set_cell_border(second_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 1, 4, 7)

    third_row = table_employment_information.rows[2]
    third_row.cells[0].text = 'Supervisor'
    third_row.cells[0].width = Inches(1.4)
    third_row.cells[1].width = Inches(2.6)
    set_cell_border(third_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 2, 2, 3)
    third_row.cells[2].text = 'Phone #'
    third_row.cells[2].width = Inches(1.4)
    set_cell_border(third_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 2, 4, 7)

    fourth_row = table_employment_information.rows[3]
    fourth_row.cells[0].text = 'Address'
    set_cell_border(fourth_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[2].text = 'City'
    set_cell_border(fourth_row.cells[3], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[4].text = 'State'
    set_cell_border(fourth_row.cells[5], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[6].text = 'ZIP'
    set_cell_border(fourth_row.cells[7], bottom={"sz": 6, "color": "#000000", "val": "single"}, )

    fifth_row = table_employment_information.rows[4]
    fifth_row.cells[0].text = 'Dates of Employment'
    fifth_row.cells[0].width = Inches(2)
    fifth_row.cells[1].width = Inches(2.6)
    set_cell_border(fifth_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 4, 2, 3)
    fifth_row.cells[2].text = 'Monthly Income $'
    fifth_row.cells[2].width = Inches(1.7)
    set_cell_border(fifth_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 4, 4, 7)

    for row in range(5):
        for cell in range(8):
            table_employment_information.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            table_employment_information.cell(row, cell).paragraphs[0].paragraph_format.space_after = Pt(0)

def get_table_employment_information_no_header():

    table_employment_information = document.add_table(rows=4, cols=8)

    employment_information = ['Current Employer', '', 'Position/Title', '']

    second_row = table_employment_information.rows[0]
    second_row.cells[0].text = 'Current Employer'
    second_row.cells[0].width = Inches(1.4)
    second_row.cells[1].width = Inches(2.6)
    set_cell_border(second_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 0, 2, 3)
    second_row.cells[2].text = 'Position/Title'
    second_row.cells[2].width = Inches(1.4)
    set_cell_border(second_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 0, 4, 7)

    third_row = table_employment_information.rows[1]
    third_row.cells[0].text = 'Supervisor'
    third_row.cells[0].width = Inches(1.4)
    third_row.cells[1].width = Inches(2.6)
    set_cell_border(third_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 1, 2, 3)
    third_row.cells[2].text = 'Phone #'
    third_row.cells[2].width = Inches(1.4)
    set_cell_border(third_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 1, 4, 7)

    fourth_row = table_employment_information.rows[2]
    fourth_row.cells[0].text = 'Address'
    set_cell_border(fourth_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[2].text = 'City'
    set_cell_border(fourth_row.cells[3], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[4].text = 'State'
    set_cell_border(fourth_row.cells[5], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    fourth_row.cells[6].text = 'ZIP'
    set_cell_border(fourth_row.cells[7], bottom={"sz": 6, "color": "#000000", "val": "single"}, )

    fifth_row = table_employment_information.rows[3]
    fifth_row.cells[0].text = 'Dates of Employment'
    fifth_row.cells[0].width = Inches(2)
    fifth_row.cells[1].width = Inches(2.6)
    set_cell_border(fifth_row.cells[1], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 3, 2, 3)
    fifth_row.cells[2].text = 'Monthly Income $'
    fifth_row.cells[2].width = Inches(1.7)
    set_cell_border(fifth_row.cells[4], bottom={"sz": 6, "color": "#000000", "val": "single"}, )
    get_merge_cells(table_employment_information, 3, 4, 7)

    for row in range(4):
        for cell in range(8):
            table_employment_information.cell(row, cell).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            table_employment_information.cell(row, cell).paragraphs[0].paragraph_format.space_after = Pt(0)



if __name__=='__main__':

    document = Document()

    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1)
        section.right_margin = Cm(1)

    get_header()
    get_header_paragraph('test')
    get_table_applicant_information()
    get_table_additional_occupant()
    get_table_residence_history()
    get_table_employment_information()
    document.add_paragraph()
    get_table_employment_information_no_header()
    document.add_paragraph()
    get_table_employment_information_no_header()


    # get_privious_address()
    # get_additionals_questions_text()
    get_table_sign()

    document.save('test1.docx')

