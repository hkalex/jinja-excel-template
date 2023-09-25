from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.chartsheet.chartsheet import Chartsheet
from openpyxl.chart import (AreaChart as PYAreaChart, Reference)
from openpyxl.utils import get_column_letter
from jinja_excel_template.helper.setxmlattr import setxmlattr
from jinja_excel_template.helper.typeconverter import str2bool

class Chart:
    """
    This Chart class is the basic chart for all the chart cell_type.
    In order you implement your own chart, you must inherit from this class.
    """

    def create_new_chart_object(self):
        raise NotImplemented()

    def set_cell(self, worksheet: Worksheet, row_num: int, column_num: int, xml_cell):
        """
        This cell type set the cell as AreaChart. row_num and column_num are 1 based index.
        """
        chart = self.create_new_chart_object()

        # all "chart." attributes will be set to the chart object
        setxmlattr(chart, xml_cell, "chart.")

        # all "x_asix." attributes will be set to the chart.x_asix object
        if hasattr(chart, "x_axis"):
            setxmlattr(chart.x_axis, xml_cell, "x_asix")

        # all "y_asix." attributes will be set to the chart.y_asix object
        if hasattr(chart, "y_axis"):
            setxmlattr(chart.y_axis, xml_cell, "y_asix")

        # set data_range
        data_range_str = xml_cell.get("data_range")
        if data_range_str == None:
            raise Exception("data_range must be provided")
        data_ref = Reference(worksheet, range_string=self._get_range_string(worksheet, data_range_str))
        
        # titles_from_data
        titles_from_data = False
        if xml_cell.get("title_from_data") != None:
            titles_from_data = str2bool(xml_cell.get("title_from_data"))
        
        chart.add_data(data_ref, titles_from_data=titles_from_data)

        # set category_range
        category_range_str = xml_cell.get("category_range")
        if category_range_str != None:
            cats_ref = Reference(worksheet, range_string=self._get_range_string(worksheet, category_range_str))
            chart.set_categories(cats_ref)

        if type(worksheet) == Chartsheet:
            worksheet.add_chart(chart)
        else:
            worksheet.add_chart(chart, get_column_letter(column_num) + str(row_num))


    def _get_range_string(self, worksheet: Worksheet, range_str:str):
        if '!' in range_str:
            return range_str
        else:
            return "'" + worksheet.title + "'!" + range_str
