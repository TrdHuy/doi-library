from dataclasses import dataclass
from typing import List, Optional, Union, Dict
from copy import deepcopy

@dataclass
class DL_Position:
    x: int
    y: int
    width: int
    height: int

@dataclass
class DL_BorderStyle:
    color: str
    width: Union[float, str]  # có thể là 1.0 hoặc "Default"
    dash_type: str

@dataclass
class DL_CellBorder:
    left: Optional[DL_BorderStyle] = None
    right: Optional[DL_BorderStyle] = None
    top: Optional[DL_BorderStyle] = None
    bottom: Optional[DL_BorderStyle] = None
    diagonal_down: Optional[DL_BorderStyle] = None
    diagonal_up: Optional[DL_BorderStyle] = None

@dataclass
class DL_Run:
    text: str
    font_name: Optional[str]
    font_size: Optional[float]
    bold: Optional[bool]
    italic: Optional[bool]
    font_color: Optional[str]
    run_index: int

@dataclass
class DL_TextParagraph:
    alignment: int
    runs: List[DL_Run]
    paragraph_index: int
    text: Optional[str] = None
    bullet: Optional[int] = None
    bullet_type: Optional[str] = None
    bullet_char: Optional[str] = None
    number_type: Optional[str] = None
    left_indent: Optional[float] = None            # Đơn vị pt
    first_line_indent: Optional[float] = None      # Đơn vị pt
    level: Optional[int] = None
    line_spacing: Optional[float] = None
    font_name:  Optional[str] = None
    font_size: Optional[float] = None 
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    font_color: Optional[str] = None

@dataclass
class DL_TextFrameFormat:
    wrap: Optional[bool]
    auto_fit: Optional[bool]
    vertical_anchor: Optional[int]
    margin: Dict[str, int]

@dataclass
class DL_Text:
    frame_format: DL_TextFrameFormat
    paragraphs: List[DL_TextParagraph]

@dataclass
class DL_MergeInfo:
    row: int
    col: int
    row_span: int
    col_span: int

@dataclass
class DL_Table:
    rows: int
    cols: int
    data: List[List[str]]
    data_detail: List[List[Optional[DL_Text]]]
    cell_fills: List[List[str]]
    merge_info: List[DL_MergeInfo]
    col_widths: List[int]
    row_heights: List[int]
    cell_borders: List[List[Optional[DL_CellBorder]]]
    
    def insert_row(
        self,
        index: int,
        content: tuple[str, ...],
        data_detail: list[Optional[DL_Text]],
        cell_fill: list[str],
        cell_border: list[Optional[DL_CellBorder]],
        row_height: int,
        merge_info: list[DL_MergeInfo]
    ):
        if len(content) != self.cols:
            raise ValueError(f"Dòng được chèn có {len(content)} phần tử, bảng có {self.cols} cột")

        if len(data_detail) != self.cols:
            raise ValueError(f"Số lượng phần tử trong data_detail không khớp số cột ({self.cols})")
        if len(cell_fill) != self.cols:
            raise ValueError(f"Số lượng phần tử trong cell_fill không khớp số cột ({self.cols})")
        if len(cell_border) != self.cols:
            raise ValueError(f"Số lượng phần tử trong cell_border không khớp số cột ({self.cols})")

        if index < 0 or index > self.rows:
            raise IndexError(f"Chỉ số dòng {index} nằm ngoài phạm vi (0 ~ {self.rows})")

        # Chèn vào các field
        self.data.insert(index, [str(cell) for cell in content])
        self.data_detail.insert(index, data_detail)
        self.cell_fills.insert(index, cell_fill)
        self.cell_borders.insert(index, cell_border)
        self.row_heights.insert(index, row_height)

        # Tăng index các merge info bên dưới (trừ merge mới)
        for mi in self.merge_info:
            if mi.row >= index:
                mi.row += 1
                
        self.merge_info.extend(deepcopy(merge_info))

        self.rows += 1
        
    def delete_row(self, index: int):
        if index < 0 or index >= self.rows:
            raise IndexError(f"Row index {index} out of range (0 ~ {self.rows-1})")

        self.data.pop(index)
        self.data_detail.pop(index)
        self.cell_fills.pop(index)
        self.cell_borders.pop(index)
        self.row_heights.pop(index)

        # Xoá và cập nhật merge_info
        new_merge_info: list[DL_MergeInfo] = []
        for m in self.merge_info:
            if m.row < index:
                new_merge_info.append(m)
            elif m.row > index:
                m.row -= 1
                new_merge_info.append(m)
            # nếu m.row == index thì bỏ merge này (vì xoá dòng đó rồi)

        self.merge_info = new_merge_info
        self.rows -= 1
        
    def build_merge_info_row(
        self,
        actual_index: int,
        template_row_index: int
    ) -> list[DL_MergeInfo]:
        template_merge_info = [mi for mi in self.merge_info if mi.row == template_row_index]
        merged_info = [deepcopy(m) for m in template_merge_info]
        for m in merged_info:
            m.row = actual_index
        return merged_info
    
    def build_data_detail_row(
        self,
        content: tuple[str, ...],
        template_row_index: Optional[int] = None
    ) -> list[Optional[DL_Text]]:
        if len(content) != self.cols:
            raise ValueError(f"Số phần tử của content ({len(content)}) không khớp với số cột ({self.cols})")

        if template_row_index is not None:
            if template_row_index < 0 or template_row_index >= self.rows:
                raise IndexError(f"template_row_index {template_row_index} vượt quá số dòng của bảng")
            template_row = self.data_detail[template_row_index]
        else:
            template_row = [None] * self.cols

        detail_row: list[Optional[DL_Text]] = []
        for col, cell_text in enumerate(content):
            template = template_row[col]
            detail = deepcopy(template)
            if detail is not None:
                # Clone style từ template
                if not detail.paragraphs:
                    raise NotImplementedError("Template DL_Text chưa hỗ trợ paragraphs rỗng")

                first_para = detail.paragraphs[0]
                if not first_para.runs:
                    raise NotImplementedError("Template DL_Text.paragraph không có run nào")

                first_run = first_para.runs[0]
                first_run.text = str(cell_text)
                first_para.runs = [first_run]
                detail.paragraphs = [first_para]
            
            detail_row.append(detail)

        return detail_row
    

@dataclass
class DL_Border:
    color: Optional[str]
    width_pt: Union[float, str, None]
    style: Optional[str]
    
@dataclass
class DL_Image:
    filename: str                  # ví dụ: "asset/img_slide1_shape3_abcd1234.png"
    ext: str                       # ví dụ: "png"
    content_type: str              # ví dụ: "image/png"
    size: int                      # kích thước byte

@dataclass
class DL_Shape:
    shape_index: int
    shape_name: str
    type: int
    position: DL_Position
    background_fill_color: Optional[str]
    border: Optional[DL_Border] = None
    text: Optional[DL_Text] = None
    table: Optional[DL_Table] = None
    image: Optional[DL_Image] = None   # 👈 Thêm dòng này
    
@dataclass
class DL_Slide:
    slide_number: int
    shapes: List[DL_Shape]
    slide_id: str
    slide_tag_info: Optional[dict[str, str]]
    
@dataclass
class DL_PPTXData:
    slide_width: int
    slide_height: int
    slides: List[DL_Slide]