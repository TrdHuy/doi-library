from abc import ABC, abstractmethod
from data.doi_template.v1.contract import *
from data.pptxdata_utils import *
from data.pptxdata import DL_Shape, DL_PPTXData, DL_Slide, DL_Text, DL_Table, DL_MergeInfo
from copy import deepcopy
from plugin.injector.injection_map.InjectionMap import InjectionMap
from typing import Callable, Any, Union
from dataclasses import dataclass

INJECT_REGISTRY: list[tuple[Callable[..., Any], Any]] = []

@dataclass
class InjectValue:
    value: Any
    # metadata kèm theo nếu cần (vd: format, style, note...)
    meta: Optional[dict[str, Any]] = None


def run_injection(pptx_data: DL_PPTXData, injection_map: InjectionMap):
    for func, injector in INJECT_REGISTRY:
        result = func(injection_map)
        if not isinstance(result, InjectValue):
            raise TypeError(
                f"Hàm {func.__name__} phải trả về InjectValue, nhưng lại nhận được {type(result)}")
        injector.inject(pptx_data, result)


class Injector(ABC):
    @abstractmethod
    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue):
        pass

    def find_slide_by_inject_id(self, pptx_data: DL_PPTXData, slide_id: str) -> Optional[DL_Slide]:
        return next((s for s in pptx_data.slides if s.slide_tag_info
                     and s.slide_tag_info.get("inject_id") == slide_id), None)

    def find_shape_by_name(self, slide: DL_Slide, shape_name: str) -> Optional[DL_Shape]:
        return next((s for s in slide.shapes if s.shape_name == shape_name), None)


class ShapeTextInjector(Injector):
    def __init__(self,
                 slide_id: SlideInjectId,
                 shape_name: ShapeName):
        self.shape_name = shape_name
        self.slide_id = slide_id

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue):
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


class TableCellInjector(Injector):
    def __init__(self,
                 slide_id: SlideInjectId,
                 table_shape_name: ShapeName,
                 cell_text_id: RunSample):
        self.table_shape_name = table_shape_name
        self.cell_text_id = cell_text_id
        self.slide_id = slide_id

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue):
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


class TableRowInserter(Injector):
    def __init__(self,
                 slide_id: SlideInjectId,
                 table_shape_name: ShapeName):
        self.slide_id = slide_id
        self.table_shape_name = table_shape_name

    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue):
        slide = self.find_slide_by_inject_id(pptx_data, self.slide_id)
        if slide is None:
            raise ValueError(
                f"Không tìm thấy slide có inject_id = {self.slide_id}")

        table_shape = self.find_shape_by_name(slide, self.table_shape_name)
        if table_shape is None or table_shape.table is None:
            raise ValueError(
                f"Không tìm thấy table shape tên = {self.table_shape_name}")

        table: DL_Table = table_shape.table
        meta = injected_value.meta
        if meta is None:
            raise ValueError("meta không được None")

        insert_index = int(meta.get("insert_index", table.rows))

        # Bắt buộc phải có template row index
        template_row_index = meta.get("template_row_index")

        if template_row_index is None:
            raise ValueError(
                "inject meta cần có 'template_row_index' để lấy template cho cell border, fill, height...")
        template_row_index = int(template_row_index)

        if template_row_index >= table.rows:
            raise IndexError(
                f"template_row_index {template_row_index} vượt quá số row hiện tại trong bảng ({table.rows})")

        row_data_list: Union[tuple[str, ...],
                             list[tuple[str, ...]]] = injected_value.value

        # Normalize
        if isinstance(row_data_list, tuple):
            row_data_list = [row_data_list]
        if isinstance(row_data_list, list):                                 # type: ignore
            if not all(isinstance(item, tuple) for item in row_data_list):  # type: ignore
                raise TypeError("Tất cả phần tử trong danh sách phải là tuple")
        else:
            raise TypeError(
                "Giá trị injected vào TableRowInserter phải là tuple hoặc list các tuple")

        for i, row_data in enumerate(row_data_list):
            if len(row_data) != table.cols:
                raise ValueError(
                    f"Tuple tại index {i} có {len(row_data)} phần tử, nhưng bảng có {table.cols} cột")

        for offset, row_data in enumerate(row_data_list):
            actual_index = insert_index + offset
            table.rows += 1
            table.data.insert(actual_index, [str(cell) for cell in row_data])
            # Clone các thuộc tính từ template
            template_detail_row = table.data_detail[template_row_index]
            template_fill_row = table.cell_fills[template_row_index]
            template_border_row = table.cell_borders[template_row_index]
            new_detail_row: list[Optional[DL_Text]] = []
            for col, cell_text in enumerate(row_data):
                detail = deepcopy(template_detail_row[col])
                if detail is not None:
                    # Clear paragraphs và chỉ giữ 1 paragraph + 1 run
                    if detail.paragraphs:
                        first_paragraph = detail.paragraphs[0]
                        if first_paragraph.runs:
                            first_run = first_paragraph.runs[0]
                            first_run.text = str(cell_text)
                            first_paragraph.runs = [first_run]
                        else:
                            raise NotImplementedError(
                                "Chưa hỗ trợ trường hợp paragraphs.runs là rỗng"
                            )
                        detail.paragraphs = [first_paragraph]
                    else:
                        raise NotImplementedError(
                            "Chưa hỗ trợ trường hợp detail.paragraphs là rỗng"
                        )
                    new_detail_row.append(detail)

            # Ghi lại data_detail đã chuẩn hoá (chỉ 1 paragraph và 1 run mỗi cell)
            table.data_detail.insert(actual_index, new_detail_row)
            table.cell_fills.insert(actual_index, list(template_fill_row))
            table.cell_borders.insert(actual_index, list(template_border_row))
            table.row_heights.insert(
                actual_index, table.row_heights[template_row_index])

            # region: Xử lý merge info
            # Merge info: nếu có merge tại hàng này thì cần update lại merge index
            # Copy merge info từ template row
            template_merge_row_info = [
                mi for mi in table.merge_info if mi.row == template_row_index]
            copied_merges: list[DL_MergeInfo] = []

            for m in template_merge_row_info:
                copied_merge = deepcopy(m)
                copied_merge.row = actual_index
                copied_merges.append(copied_merge)
                table.merge_info.append(copied_merge)

            # Dùng id để tránh tăng row cho những merge vừa tạo hoặc là của template
            excluded_ids = {id(m) for m in copied_merges}

            for mi in table.merge_info:
                if mi.row >= actual_index and id(mi) not in excluded_ids:
                    mi.row += 1

            # endregion

            # Ghi lại text thực tế (chèn nội dung mới nhưng giữ nguyên format từ template)
            for col, cell_text in enumerate(row_data):
                detail = deepcopy(template_detail_row[col])
                if detail is not None:
                    # Ghi đè text vào run đầu tiên
                    if detail.paragraphs and detail.paragraphs[0].runs:
                        detail.paragraphs[0].runs[0].text = str(cell_text)
                table.data_detail[actual_index][col] = detail

            print(f"✔️ Đã chèn dòng vào index {actual_index}: {row_data}")
