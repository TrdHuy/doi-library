from data.doi_template.v1.basic_info import BasicInfo
from plugin.injector.machine.injector_base import *
from .decorator import inject_with
from dataclasses import dataclass


class BasicInfoInjectionMap:
    def __init__(self, basic_info: BasicInfo):
        self.basic_info = basic_info

    @inject_with(ShapeTextInjector(SlideInjectId.TITLE_SLIDE, ShapeName.TITLE))
    def get_title(self):
        return InjectValue(value=self.basic_info.title)

    @inject_with(TableCellInjector(SlideInjectId.BASIC_INFO_SLIDE, ShapeName.BASIC_INFO_TABLE, RunSample.DEPARTMENT_RS))
    def get_department(self):
        return InjectValue(value=self.basic_info.department)

    @inject_with(TableCellInjector(SlideInjectId.BASIC_INFO_SLIDE, ShapeName.BASIC_INFO_TABLE, RunSample.PROJECT_NAME_RS))
    def get_project_name(self):
        return InjectValue(value=self.basic_info.project_name)

    @inject_with(TableRowInserter(SlideInjectId.BASIC_INFO_SLIDE, ShapeName.BASIC_INFO_TABLE))
    def get_inventors(self):
        inventor_rows = [
            (inv.no, inv.full_name, inv.contribution,
             inv.employee_no, "empty_cell", inv.status)
            for inv in self.basic_info.inventors
        ]
        return InjectValue(
            value=inventor_rows,
            meta={
                "insert_index": 5,
                "template_row_index": 5,  # Dòng mẫu đã có định dạng sẵn
            }
        )
