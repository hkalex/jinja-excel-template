from jinja_excel_template.helper.GetFill import get_fill
from lxml import etree

def test_get_font():
    xml_string = """<cell value="1-A1"
fill.fill_type="solid"
fill.start_color="FFFFFFFF"
fill.end_color="00000000"
/>
"""
    xml_cell = etree.fromstring(xml_string)
    fill = get_fill(xml_cell)
    assert fill.fill_type == "solid"
    assert fill.start_color.rgb == "FFFFFFFF"
    assert fill.end_color.rgb == "00000000"
