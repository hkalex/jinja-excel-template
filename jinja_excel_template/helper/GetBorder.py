from openpyxl.styles import Border, Side
from jinja_excel_template.helper.setxmlattr import setxmlattr

def get_border(xml_cell):

# >>> border = Border(left=Side(border_style=None,
# ...                           color='FF000000'),
# ...                 right=Side(border_style=None,
# ...                            color='FF000000'),
# ...                 top=Side(border_style=None,
# ...                          color='FF000000'),
# ...                 bottom=Side(border_style=None,
# ...                             color='FF000000'),
# ...                 diagonal=Side(border_style=None,
# ...                               color='FF000000'),
# ...                 diagonal_direction=0,
# ...                 outline=Side(border_style=None,
# ...                              color='FF000000'),
# ...                 vertical=Side(border_style=None,
# ...                               color='FF000000'),
# ...                 horizontal=Side(border_style=None,
# ...                                color='FF000000')
# ...                )

    prefix = "border."

    border = Border()

    # set all the other non-Side attributes, like diagonal_direction
    setxmlattr(border, xml_cell, prefix)

    for side_name in ['left', 'right', 'top', 'bottom', 'diagonal', 'outline', 'vertical', 'horizontal']:
        side = Side()
        for attr_name in ['border_style', 'color']:
            attr_value = xml_cell.get(prefix + side_name + '.' + attr_name)
            if attr_value != None: setattr(side, attr_name, attr_value)
        setattr(border, side_name, side)

    return border
