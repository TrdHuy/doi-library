from abc import ABC, abstractmethod
from data.doi_template.v1.contract import *
from data.pptxdata_utils import *
from data.pptxdata import DL_Shape, DL_PPTXData, DL_Slide, DL_Table
from plugin.injector.injection_map.InjectionMap import InjectionMap
from typing import Callable, Any, Union
from dataclasses import dataclass,  field

INJECT_REGISTRY: list[tuple[Callable[..., Any], Any]] = []


class InjectMetaKey(str, Enum):
    INSERT_INDEX = "INSERT_INDEX"
    TEMPLATE_ROW_INDEX = "TEMPLATE_ROW_INDEX"
    IS_DELETE_TEMPLATE_ROW = "IS_DELETE_TEMPLATE_ROW"


@dataclass
class InjectValue:
    value: Any
    __meta: dict[str, Any] = field(
        default_factory=dict[str, Any], init=False, repr=False)

    def __init__(self, value: Any, meta: Optional[dict[InjectMetaKey, Any]] = None):
        self.value = value
        self.__meta = {k.value: v for k, v in meta.items()} if meta else {}

    def __getitem__(self, key: InjectMetaKey) -> Optional[Any]:
        return self.__meta.get(key.value)

    def __setitem__(self, key: InjectMetaKey, val: Any) -> None:
        self.__meta[key.value] = val

    def __contains__(self, key: InjectMetaKey) -> bool:
        return key.value in self.__meta

    def remove(self, key: InjectMetaKey) -> None:
        self.__meta.pop(key.value, None)

    def keys(self):
        return [InjectMetaKey(k) for k in self.__meta]

    def get(self, key: InjectMetaKey, default: Any = None) -> Any:
        return self.__meta.get(key.value, default)

    def get_int(self, key: InjectMetaKey, default: int = 0) -> int:
        val = self.get(key, default)
        try:
            return int(val)
        except (ValueError, TypeError):
            return default

    def __repr__(self):
        return f"InjectValue(value={self.value})"


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
