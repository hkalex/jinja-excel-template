from .ExcelXMLMerger import ExcelXMLMerger
from .ExcelGenerator import ExcelGenerator

class JinjaExcelEngine:

    options = {}

    def generate(self, data, output_file_path, jinja_template_str:str = None, jinja_template_path:str = None):
        merger = ExcelXMLMerger()
        xml = merger.merge(data, template_string=jinja_template_str, template_file_path=jinja_template_path)

        generator = ExcelGenerator()
        generator.generate_excel(output_file_path, xml)

