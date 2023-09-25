from jinja_excel_template.ExcelGenerator import ExcelGenerator
import os

dir_path = os.path.dirname(os.path.realpath(__file__))
print(dir_path)

def test_both_none():
    try:
        generator = ExcelGenerator()
        generator.generate_excel()
    except Exception:
        # all good
        pass

def test_xml_file_path():
    generator = ExcelGenerator()
    generator.generate_excel(dir_path + "/temp/test1.xlsx", xml_file_path = dir_path + "/test.xml")

def test_xml_string():
    f = open(dir_path + "/test.xml")
    xml_string = f.read()
    generator = ExcelGenerator()
    generator.generate_excel(dir_path + "/temp/test2.xlsx", xml_string=xml_string)
