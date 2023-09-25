from openpyxl.worksheet.worksheet import Worksheet
from dateutil.parser import parse
from jinja_excel_template.helper.utils import is_date


class DateTime:
    def set_cell(self, worksheet: Worksheet, row: int, column: int, xml_cell):
        """
        This cell type set the cell as general cell. row and column are 1 based index.
        """

        if xml_cell == None: return


        value_str = xml_cell.get("value")

        if is_date(value_str):
            dt = parse(value_str)
            worksheet.cell(row, column, dt)
        else:
            worksheet.cell(row, column, value_str)
