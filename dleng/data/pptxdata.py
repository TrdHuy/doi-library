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