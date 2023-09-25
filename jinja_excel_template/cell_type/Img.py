from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from urllib import request
import uuid
import os
import tempfile

class Img:
    def set_cell(self, worksheet:Worksheet, row:int, column:int, xml_cell):
        """
        This cell type set the cell as general cell. row and column are 1 based index.
        """

        url = xml_cell.get("url")
        temp_file_name = os.path.join(tempfile.gettempdir(), str(uuid.uuid1()))
        request.urlretrieve(url, temp_file_name)

        column_letter = get_column_letter(column)

        img = Image(temp_file_name)
        if xml_cell.get("width") != None: img.width = float(xml_cell.get("width"))
        if xml_cell.get("height") != None: img.height = float(xml_cell.get("height"))

        worksheet.add_image(img, column_letter + str(row))

