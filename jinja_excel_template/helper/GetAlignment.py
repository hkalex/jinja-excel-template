from openpyxl.styles import Alignment
from jinja_excel_template.helper.setxmlattr import setxmlattr

def get_alignment(xml_cell):
    alignment = Alignment()
    setxmlattr(alignment, xml_cell, "align.")
    return alignment