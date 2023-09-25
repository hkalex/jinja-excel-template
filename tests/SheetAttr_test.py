from jinja_excel_template.helper.setSheetAttr import set_sheet_attr
from lxml import etree
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook



def test_set_sheet_attr():
    xml_string = """<sheet
    print_title_cols="A:A"
    print_title_rows="1:1"
    print_area="A1:B2"
    page_setup.orientation="landscape"
    page_setup.paperSize="11"
    />
    """

    xml_sheet = etree.fromstring(xml_string)

    wb = Workbook()
    ws = wb.active

    set_sheet_attr(ws, xml_sheet)

    assert ws.page_setup.orientation == "landscape"
    assert ws.print_title_cols == "$A:$A"
    assert ws.print_title_rows == "$1:$1"
    assert ws.print_area == "'Sheet'!$A$1:$B$2"
    assert ws.page_setup.paperSize == 11