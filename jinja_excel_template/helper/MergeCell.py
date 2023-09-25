from openpyxl.worksheet.worksheet import Worksheet

def merge_cell(ws:Worksheet, row_number, column_number, xml_cell):
    merge_right = xml_cell.get("merge_right")
    merge_down = xml_cell.get("merge_down")

    merge_right_num = 0
    merge_down_num = 0

    if merge_right != None and str.isnumeric(merge_right): merge_right_num = int(merge_right)
    if merge_down != None and str.isnumeric(merge_down): merge_down_num = int(merge_down)

    if merge_right_num > 0 or merge_down_num > 0:
        ws.merge_cells(start_row = row_number, end_row= row_number+merge_down_num, start_column=column_number, end_column=column_number+merge_right_num)
