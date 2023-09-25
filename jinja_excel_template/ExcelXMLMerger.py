import jinja2
import os


class ExcelXMLMerger:
    """
    This class merges data and jinja template
    """

    options = {}
    templateLoader = jinja2.FileSystemLoader( "/" )
    templateEnv = jinja2.Environment( loader=templateLoader )

    def __init__(self, template_folder = None):
        if template_folder != None:
            self.templateLoader = jinja2.FileSystemLoader( template_folder )
            self.templateEnv = jinja2.Environment( loader=self.templateLoader )


    def merge(self, data, template_string = None, template_file_path = None):
        if template_string != None:
            template = self.templateEnv.from_string(template_string)
        elif template_file_path != None:
            template = self.templateEnv.get_template(template_file_path)
        else:
            raise Exception("template_string and template_file_path cannot be both None")
        
        result = template.render(data)
        return result
