from os import path
import json
from jinja_excel_template.JinjaExcelEngine import JinjaExcelEngine

def test_countries():
    data = prepare_data()

    output_file_path = path.join(path.dirname(__file__), "temp", "countries.xlsx")
    template_file_path = path.join(path.dirname(__file__), "countries.jinja")

    engine = JinjaExcelEngine()
    engine.generate(data, output_file_path, jinja_template_path=template_file_path)
    
    


def count_by(arr, field_name):
    result = {}
    for ele in arr:
        if field_name in ele:
            value = ele[field_name]
            if value in result:
                result[value] += 1
            else:
                result[value] = 1
    return result


def prepare_data():
    """
    I didn't use any external library to prepare the data, because I want to keep this
    package to be as small as possible. You can use numpy or pandas if you want.

    The final data looks like this
    {
        "regions": {'AS': 51, 'EU': 47, 'AF': 57, 'OC': 25, 'NA': 35, 'SA': 16, 'Americas': 2},
        "details": {
            "AS": [
                {
                    ... list of country in AS
                }
            ], ... and other regions
        }
    }
    """
    countries_file_path = path.join(path.dirname(__file__), "countries.json")
    f = open(countries_file_path, 'r')
    countries = json.load(f)

    regions = count_by(countries, 'region')
    details = {}
    for key in regions.keys():
        details[key] = list(filter(lambda c: c['region'] == key, countries ))

    return {
        "regions": regions,
        "details": details
    }
