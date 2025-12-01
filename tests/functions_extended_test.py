from jinja_excel_template.formula_eval import safe_eval
from datetime import datetime


def test_round_int_abs_min_max():
    assert safe_eval("ROUND(1.234,2)") == 1.23
    assert safe_eval("INT(3.9)") == 3
    assert safe_eval("ABS(-5)") == 5
    assert safe_eval("MAX([1,2,3])") == 3
    assert safe_eval("MIN([1,2,3])") == 1


def test_date_fn():
    d1 = safe_eval("DATE('2020-01-10')")
    assert isinstance(d1, datetime)
    d2 = safe_eval("DATE(2020,1,10)")
    assert isinstance(d2, datetime)
