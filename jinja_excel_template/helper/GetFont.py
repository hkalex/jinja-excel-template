from openpyxl.styles import Font
from jinja_excel_template.helper.setxmlattr import setxmlattr


def get_font(xml_cell):
#     font = Font(name='Calibri',
# ...                 size=11,
# ...                 bold=False,
# ...                 italic=False,
# ...                 vertAlign=None,
# ...                 underline='none',
# ...                 strike=False,
# ...                 color='FF000000')
    font = Font()
    
    setxmlattr(font, xml_cell, "font.")

    return font
