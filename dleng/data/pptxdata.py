from dataclasses import dataclass, field
from typing import List, Optional, Union, Dict, Any

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
class DL_Paragraph:
    alignment: Optional[int]
    runs: List[DL_Run]
    paragraph_index: int

@dataclass
class DL_TextParagraph:
    alignment: Optional[int]
    runs: List[DL_Run]
    paragraph_index: int
    text: Optional[str] = None

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
    data_detail: List[List[Optional[List[DL_Paragraph]]]]
    cell_fills: List[List[str]]
    merge_info: List[DL_MergeInfo]
    col_widths: List[int]
    row_heights: List[int]
    cell_borders: List[List[Optional[DL_CellBorder]]]

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
    type: int
    position: DL_Position
    background_fill_color: Optional[str]
    border: Optional[DL_Border] = None
    text: Optional[List[DL_TextParagraph]] = None
    table: Optional[DL_Table] = None
    image: Optional[DL_Image] = None   # 👈 Thêm dòng này

@dataclass
class DL_Slide:
    slide_number: int
    shapes: List[DL_Shape]

@dataclass
class DL_PPTXData:
    slide_width: int
    slide_height: int
    slides: List[DL_Slide]
