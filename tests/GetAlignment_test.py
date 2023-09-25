from jinja_excel_template.helper.GetAlignment import get_alignment
from lxml import etree

def test_get_alignment():
    xml_string = """<cell value="1-A1"
    align.horizontal="right"
    align.vertical="top"
    align.wrap_text="true"
    align.shrink_to_fit="true"
    align.indent="10"
/>
"""
    xml_cell = etree.fromstring(xml_string)
    alignment = get_alignment(xml_cell)

    assert alignment.horizontal == "right"
    assert alignment.vertical == "top"
    assert alignment.wrap_text == True
    assert alignment.shrink_to_fit == True
    assert alignment.indent == 10

    

