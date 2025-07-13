from abc import ABC, abstractmethod
from data.doi_template.v1.contracts.pptx_contract import *
from data.pptxdata_utils import *
from data.pptxdata import DL_Shape, DL_PPTXData, DL_Slide, DL_Table
from plugin.injector.injection_map.InjectionMap import InjectionMap
from typing import Callable, Any, Union
from data.doi_template.v1.slide_data_model import SectionContent, SlideContent, SlideBlock
from typing import Dict, List, Tuple, Type, Generic, Optional, Any
from .InjectValue import T, InjectValue, InjectMetaKey


class Injector(ABC, Generic[T]):
    @abstractmethod
    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue[T]) -> None:
        pass

    def find_slide_by_inject_id(self, pptx_data: DL_PPTXData, slide_id: str) -> Optional[DL_Slide]:
        return next((s for s in pptx_data.slides if s.slide_tag_info
                     and s.slide_tag_info.get("inject_id") == slide_id), None)

    def find_shape_by_name(self, slide: DL_Slide, shape_name: str) -> Optional[DL_Shape]:
        return next((s for s in slide.shapes if s.shape_name == shape_name), None)

class TemplateShapeTextInjector(Injector[str]):
    def __init__(self,
                 slide_id: SlideInjectId,
                 shape_name: ShapeName):
        self.shape_name = shape_name
        self.slide_id = slide_id

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue[str]):
        slide = self.find_slide_by_inject_id(pptx_data, self.slide_id)
        if slide is None:
            raise ValueError(
                f"Không tìm thấy slide có inject_id = {self.slide_id}")

        shape = self.find_shape_by_name(slide, self.shape_name)
        if shape is None:
            raise ValueError(
                f"Không tìm thấy table shape có tên = {self.shape_name}")
        if shape.text is None:
            raise NotImplementedError(
                "Chưa hỗ trợ nếu shape không có paragraph/run mẫu")

        paragraphs = shape.text.paragraphs
        if not paragraphs or not paragraphs[0].runs or not paragraphs[0].runs[0]:
            raise NotImplementedError(
                "Chưa hỗ trợ nếu shape không có paragraph/run mẫu")

        template_run = paragraphs[0].runs[0]
        new_run = make_run_from_template(template_run, injected_value.value)
        shape.text.paragraphs[0].runs[0] = new_run

class TableCellInjector(Injector[str]):
    def __init__(self,
                 slide_id: SlideInjectId,
                 table_shape_name: ShapeName,
                 cell_text_id: RunSample):
        self.table_shape_name = table_shape_name
        self.cell_text_id = cell_text_id
        self.slide_id = slide_id

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue[str]):
        slide = self.find_slide_by_inject_id(pptx_data, self.slide_id)
        if slide is None:
            raise ValueError(
                f"Không tìm thấy slide có inject_id = {self.slide_id}")

        table_shape = self.find_shape_by_name(slide, self.table_shape_name)
        if table_shape is None or table_shape.table is None:
            raise ValueError(
                f"Không tìm thấy table shape có tên = {self.table_shape_name}")

        idx = find_table_cell_by_text(table_shape.table, self.cell_text_id)
        if idx:
            row, col = idx
            print(f"Found at ({row}, {col})")
            set_table_cell_text(table_shape.table, row,
                                col, injected_value.value)
        else:
            raise NotImplementedError(
                "Chưa hỗ trợ nếu cell không chứa " + self.cell_text_id)

class TableRowInserter(Injector[list[tuple[str]]]):
    def __init__(self,
                 slide_id: SlideInjectId,
                 table_shape_name: ShapeName):
        self.slide_id = slide_id
        self.table_shape_name = table_shape_name

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue[list[tuple[str]]]):
        slide = self.find_slide_by_inject_id(pptx_data, self.slide_id)
        if slide is None:
            raise ValueError(
                f"Không tìm thấy slide có inject_id = {self.slide_id}")

        table_shape = self.find_shape_by_name(slide, self.table_shape_name)
        if table_shape is None or table_shape.table is None:
            raise ValueError(
                f"Không tìm thấy table shape tên = {self.table_shape_name}")

        table: DL_Table = table_shape.table

        insert_index = injected_value.get_int(
            InjectMetaKey.INSERT_INDEX, table.rows)

        # Bắt buộc phải có template row index
        template_row_index = injected_value[InjectMetaKey.TEMPLATE_ROW_INDEX]

        if template_row_index is None:
            raise ValueError(
                "inject meta cần có 'template_row_index' để lấy template cho cell border, fill, height...")
        template_row_index = int(template_row_index)

        if template_row_index >= table.rows:
            raise IndexError(
                f"template_row_index {template_row_index} vượt quá số row hiện tại trong bảng ({table.rows})")

        row_data_list: list[tuple[str]] = injected_value.value
        if not all(isinstance(item, tuple) for item in row_data_list):
                raise TypeError("Tất cả phần tử trong danh sách phải là tuple")
       

        for i, row_data in enumerate(row_data_list):
            if len(row_data) != table.cols:
                raise ValueError(
                    f"Tuple tại index {i} có {len(row_data)} phần tử, nhưng bảng có {table.cols} cột")

        # Dữ liệu từ template row
        template_fill = table.cell_fills[template_row_index]
        template_border = table.cell_borders[template_row_index]
        template_height = table.row_heights[template_row_index]

        for offset, row_data in enumerate(row_data_list):
            if len(row_data) != table.cols:
                raise ValueError(
                    f"Tuple tại index {offset} có {len(row_data)} phần tử, bảng có {table.cols} cột")

            actual_index = insert_index + offset

            # Gọi insert
            table.insert_row(
                index=actual_index,
                content=row_data,
                data_detail=table.build_data_detail_row(
                    row_data, template_row_index),
                cell_fill=list(template_fill),
                cell_border=list(template_border),
                row_height=template_height,
                merge_info=table.build_merge_info_row(
                    actual_index=actual_index,
                    template_row_index=template_row_index)
            )

            print(f"✔️ Đã chèn dòng vào index {actual_index}: {row_data}")
        num_rows_inserted_before_template = 0 if template_row_index < insert_index else len(
            row_data_list)
        adjusted_template_row_index = template_row_index + \
            num_rows_inserted_before_template
        if injected_value.get(InjectMetaKey.IS_DELETE_TEMPLATE_ROW, False):
            table.delete_row(adjusted_template_row_index)
            print(
                f"🗑️ Đã xoá dòng template tại index {adjusted_template_row_index}")


INJECT_REGISTRY: Dict[
    Type[InjectionMap],
    List[Tuple[Callable[..., InjectValue[Any]], Injector[Any]]]
] = {}


def run_injection(pptx_data: DL_PPTXData, injection_map: InjectionMap):
    cls = type(injection_map)
    inject_list = INJECT_REGISTRY.get(cls, [])

    for func, injector in inject_list:
        result = func(injection_map)
        injector.inject(pptx_data, result)


# class SectionContentInjector(Injector):
#     def __init__(self, slide_id: SlideInjectId,
#                  text_content_shape_name: ShapeName,
#                  image_content_shape_name: ShapeName):
#         self.slide_id = slide_id
#         self.text_content_shape_name = text_content_shape_name
#         self.image_content_shape_name = image_content_shape_name

#     def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue) -> None:
#         section: SectionContent = injected_value.value

#         for slide_index, slide_content in enumerate(section.slides):
#             if slide_index >= len(slide_content.blocks):
#                 continue

#             for block in slide_content.blocks:
#                 dl_text = self.build_dl_text(block)
#                 block_value = InjectValue(value=dl_text)

#                 injector = TemplateShapeTextInjector(
#                     slide_id=self.slide_id,
#                     shape_name=self.text_content_shape_name
#                 )
#                 injector.inject(pptx_data, block_value)

#     def build_dl_text(self, block: SlideBlock) -> DL_Text:
#         lines = []

#         if block.heading:
#             lines.append(DL_TextLine(text=block.heading, is_bold=True))

#         for item in block.items:
#             lines.append(
#                 DL_TextLine(
#                     text=item.text,
#                     is_bullet=(item.type == "bullet"),
#                     bullet_level=item.level if item.type == "bullet" else 0
#                 )
#             )

#         paragraph = DL_TextParagraph(lines=lines)
#         return DL_Text(paragraphs=[paragraph])