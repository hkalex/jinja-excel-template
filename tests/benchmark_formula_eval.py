import os
from jinja_excel_template.formula_eval import safe_eval


def make_expr(n):
    return f"1 + {n}"


def test_safe_eval_benchmark(benchmark):
    # microbenchmark: evaluate 1000 tiny expressions
    expressions = [make_expr(i) for i in range(1000)]

    def run_all():
        for e in expressions:
            safe_eval(e)

    benchmark(run_all)


def test_generator_benchmark(benchmark):
    # generate an xml with 1000 rows with a formula each and measure generation.
    xml_parts = ["<Excel><sheet sheet_name=\"Bench\">"]
    for i in range(1000):
        xml_parts.append(f"<row><cell type=\"Number\" formula=\"{i} + 1\"/></row>")
    xml_parts.append("</sheet></Excel>")
    xml = "".join(xml_parts)
    from jinja_excel_template.ExcelGenerator import ExcelGenerator
    out = os.path.join(os.path.dirname(__file__), "temp", "bench.xlsx")
    generator = ExcelGenerator()

    def run():
        generator.generate_excel(out, xml_string=xml)

    benchmark(run)
