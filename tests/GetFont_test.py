from jinja_excel_template.helper.GetFont import get_font
from lxml import etree

def test_get_font():
    xml_string = """<cell value="1-A1"
font.name="Calibri"
font.size="20"
font.bold="true"
font.italic="true"
font.vertAlign="superscript"
font.underline="double"
font.strike="true"
font.color="FF000000" 
/>
"""
    xml_cell = etree.fromstring(xml_string)
    font = get_font(xml_cell)
    assert font.color.rgb == "FF000000"
    assert font.name == "Calibri"
    assert font.size == 20
    assert font.bold == True
    assert font.italic == True
    assert font.vertAlign == "superscript"
    assert font.underline == "double"
    assert font.strike == True
    assert font.color.rgb == "FF000000"