from openpyxl.worksheet.worksheet import Worksheet
from jinja_excel_template.helper.GetFont import get_font


class General:
    def set_cell(self, worksheet: Worksheet, row: int, column: int, xml_cell):
        """
        This cell type set the cell as general cell. row and column are 1 based index.
        """
        if xml_cell != None:
            value = xml_cell.get("value")
            if value: worksheet.cell(row, column, value)

