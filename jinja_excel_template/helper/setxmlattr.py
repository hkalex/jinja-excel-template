from jinja_excel_template.helper.typeconverter import str2bool
import inspect

def setxmlattr(obj, xml_node, prefix:str):
    """
    This is a simple method to dynamically set the xml attributes to the object.
    Only bool, int, float and string will be applied.
    """
    for (name, value) in inspect.getmembers(obj):
        attr_value = xml_node.get(prefix + name)
        if attr_value != None:
            if type(value) == bool:
                setattr(obj, name, str2bool(attr_value))
            elif type(value) == int:
                setattr(obj, name, int(attr_value))
            elif type(value) == float:
                setattr(obj, name, float(attr_value))
            else:
                setattr(obj, name, attr_value)