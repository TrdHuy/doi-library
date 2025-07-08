from dataclasses import dataclass
from typing import List, Optional, Literal, Union

# Loại block có thể hiển thị được
ContentBlockType = Literal["title", "paragraph", "bullet", "image", "table"]


@dataclass
class ContentBlock:
    type: ContentBlockType  # "paragraph", "bullet", "image", "title", "table"
    text: Optional[str] = None
    level: Optional[int] = 0
    image_path: Optional[str] = None
    bullet_char: Optional[str] = None
    config: Optional[dict] = None
    table_data: Optional[List[List[str]]] = None  

@dataclass
class BackgroundSlide:
    id: str
    title: Optional[str] = None
    blocks: List[ContentBlock] = None
    config: Optional[dict] = None


@dataclass
class BackgroundSection:
    name: str                              # ví dụ: "technical_field", "problem", "prior_art"
    slides: List[BackgroundSlide]          # danh sách các slide con


@dataclass
class Background:
    sections: List[BackgroundSection]      # toàn bộ các mục trong background
