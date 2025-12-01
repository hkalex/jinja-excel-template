from openpyxl.worksheet.worksheet import Worksheet
from jinja_excel_template.helper.GetFont import get_font
from jinja_excel_template.helper.evaluate_formula import evaluate_formula

from jinja_excel_template.helper.utils import is_float


class Number:
    def set_cell(self, worksheet:Worksheet, row:int, column:int, xml_cell):
        """
        This cell type set the cell as general cell. row and column are 1 based index.
        """
        if xml_cell is None:
            return

        formula = xml_cell.get("formula")
        pass_through = xml_cell.get("pass_through") or "false"
        evaluate = xml_cell.get("evaluate") or "false"

        if formula:
            # default: write Excel formula; if evaluate=true do server-side evaluation
            if not (evaluate.lower() in ("true", "1")):
                formula_text = formula if formula.startswith("=") else f"={formula}"
                worksheet.cell(row, column, formula_text)
                return
            try:
                ctx = {}
                # Context mapping: all prior cells in this sheet up to current row
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
            str_value = str(value) if value is not None else None
        else:
            str_value = xml_cell.get("value")

        if str_value is None:
            worksheet.cell(row, column, None)
            return

        if is_float(str_value):
            float_value = float(str_value)
            if float_value.is_integer():
                worksheet.cell(row, column, int(str_value))
            else:
                worksheet.cell(row, column, float_value)
        else:
            worksheet.cell(row, column, str_value)

