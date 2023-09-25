from openpyxl.worksheet.worksheet import Worksheet

sheet_attr = {
    "print_title_cols": str,
    "print_title_rows": str,
    "print_area": str
}

page_setup_attr = {
    "page_setup.orientation": str,
    "page_setup.paperSize": int
}

def set_sheet_attr(worksheet:Worksheet, xml_sheet):
    """
    This function set the <sheet> attributes to the Excel  worksheet
    """

    for attr in sheet_attr:
        attr_value = xml_sheet.get(attr)
        if attr_value != None:
            setattr(worksheet, attr, sheet_attr.get(attr)(attr_value))
    
    # set page_setup
    for attr in page_setup_attr:
        attr_value = xml_sheet.get(attr)
        if attr_value != None:
            cut_attr_name = attr.replace('page_setup.', '')
            setattr(worksheet.page_setup, cut_attr_name, page_setup_attr.get(attr)(attr_value) )

