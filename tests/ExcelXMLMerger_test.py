from jinja_excel_template.ExcelXMLMerger import ExcelXMLMerger
import os

dir_path = os.path.dirname(os.path.realpath(__file__))

def test_both_null():
    try:
        merger = ExcelXMLMerger()
        result = merger.merge(data={})
    except Exception:
        # all good
        pass

def test_template_string():
    merger = ExcelXMLMerger()
    result = merger.merge({ "Hello": "World" }, template_string="This is {{Hello}}!!")
    assert result == "This is World!!"

def test_template_file():
    merger = ExcelXMLMerger(dir_path)
    result = merger.merge({ "Hello": "World" }, template_file_path="test.jinja")
    assert result == "This is World!!"