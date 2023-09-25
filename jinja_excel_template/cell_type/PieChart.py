from openpyxl.chart import (PieChart as OpenPyXLChart)
from jinja_excel_template.cell_type.Chart import Chart

class PieChart(Chart):
    def create_new_chart_object(self):
        return OpenPyXLChart()
