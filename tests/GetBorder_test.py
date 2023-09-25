from jinja_excel_template.helper.GetBorder import get_border
from lxml import etree

def test_get_font():
    xml_string = """<cell value="1-A1"
    border.left.border_style="thick"
    border.left.color="00FF0000"

    border.right.border_style="thick"
    border.right.color="00FF0000"

    border.top.border_style="thick"
    border.top.color="00FF0000"

    border.bottom.border_style="thick"
    border.bottom.color="00FF0000"

    border.diagonal.border_style="thick"
    border.diagonal.color="00FF0000"

    border.outline="True"

    border.vertical.border_style="thick"
    border.vertical.color="00FF0000"

    border.horizontal.border_style="thick"
    border.horizontal.color="00FF0000"
/>
"""
    xml_cell = etree.fromstring(xml_string)
    border = get_border(xml_cell)

    assert border.left.border_style == "thick"
    assert border.left.color.rgb == "00FF0000"

    assert border.right.border_style == "thick"
    assert border.right.color.rgb == "00FF0000"

    assert border.top.border_style == "thick"
    assert border.top.color.rgb == "00FF0000"

    assert border.bottom.border_style == "thick"
    assert border.bottom.color.rgb == "00FF0000"

    assert border.diagonal.border_style == "thick"
    assert border.diagonal.color.rgb == "00FF0000"

    assert border.outline == True

    assert border.vertical.border_style == "thick"
    assert border.vertical.color.rgb == "00FF0000"

    assert border.horizontal.border_style == "thick"
    assert border.horizontal.color.rgb == "00FF0000"

