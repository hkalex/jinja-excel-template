from openpyxl.chart import (AreaChart as PYAreaChart)
from jinja_excel_template.cell_type.Chart import Chart

class AreaChart(Chart):
    def create_new_chart_object(self):
        return PYAreaChart()
