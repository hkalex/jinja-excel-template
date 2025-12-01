from openpyxl.worksheet.worksheet import Worksheet
from dateutil.parser import parse
from jinja_excel_template.helper.utils import is_date
from jinja_excel_template.helper.evaluate_formula import evaluate_formula


class DateTime:
    def set_cell(self, worksheet: Worksheet, row: int, column: int, xml_cell):
        """
        This cell type set the cell as general cell. row and column are 1 based index.
        """

        if xml_cell is None: return

        formula = xml_cell.get("formula")
        evaluate = xml_cell.get("evaluate") or "false"

        if formula:
            if not (evaluate.lower() in ("true", "1")):
                formula_text = formula if formula.startswith("=") else f"={formula}"
                worksheet.cell(row, column, formula_text)
                return
            try:
                ctx = {}
                max_col = worksheet.max_column or column
                max_row = row
                from openpyxl.utils import get_column_letter
                for r in range(1, max_row + 1):
                    for c in range(1, max_col + 1):
                        addr = f"{get_column_letter(c)}{r}"
                        ctx[addr] = worksheet.cell(r, c).value

                def range_resolver(start, end, sheet_name=None):
                    from openpyxl.utils import column_index_from_string
                    from openpyxl.utils.cell import coordinate_from_string
                    if sheet_name:
                        ws = worksheet.parent[sheet_name]
                    else:
                        ws = worksheet
                    sc, sr = coordinate_from_string(start)
                    ec, er = coordinate_from_string(end)
                    sc_i = column_index_from_string(sc)
                    ec_i = column_index_from_string(ec)
                    results = []
                    for rr in range(sr, er + 1):
                        for cc in range(sc_i, ec_i + 1):
                            addr = f"{get_column_letter(cc)}{rr}"
                            results.append(ws.cell(rr, cc).value)
                    return results

                ctx['__range_resolver__'] = range_resolver
                value = evaluate_formula(formula, context=ctx)
            except Exception:
                worksheet.cell(row, column, "#ERROR")
                return
            if is_date(value):
                dt = parse(value)
                worksheet.cell(row, column, dt)
            else:
                worksheet.cell(row, column, value)
            return

        value_str = xml_cell.get("value")
        if is_date(value_str):
            dt = parse(value_str)
            worksheet.cell(row, column, dt)
        else:
            worksheet.cell(row, column, value_str)
