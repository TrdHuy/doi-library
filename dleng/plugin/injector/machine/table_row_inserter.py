from .injector_base import Injector
from .inject_value import InjectValue, InjectMetaKey
from data.doi_template.v1.contracts.pptx_contract import SlideInjectId, ShapeName
from data.pptxdata import DL_PPTXData, DL_Table


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
