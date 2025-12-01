from jinja_excel_template.formula_eval import safe_eval, EvaluationError


def test_arithmetic_simple():
    assert safe_eval("1 + 2 * 3") == 7
    assert safe_eval("(1 + 2) * 3") == 9


def test_functions_sum_average_concat():
    assert safe_eval("SUM([1,2,3])") == 6
    assert safe_eval("AVERAGE([2,4,6])") == 4
    assert safe_eval("CONCAT(['a','b','c'])") == 'abc'


def test_variables_in_context():
    ctx = {"x": 10, "y": 5}
    assert safe_eval("x + y", context=ctx) == 15


def test_prohibited_exec():
    try:
        _ = safe_eval("__import__('os').system('echo hello')")
        assert False, "should not allow imports"
    except EvaluationError:
        pass
