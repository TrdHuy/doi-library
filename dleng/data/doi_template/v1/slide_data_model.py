from dataclasses import dataclass, field
from typing import List, Optional
import os

@dataclass
class SlideBlockItem:
    text: str
    type: str  # "paragraph" | "bullet"
    level: int = 0  # Với bullet: 0, 1, 2; với paragraph thì luôn là 0
    
@dataclass
class SlideBlock:
    heading: Optional[str] = None
    items: List[SlideBlockItem] = field(default_factory=list[SlideBlockItem])
    image_path: Optional[str] = None
    layout_hint: Optional[str] = None           # Gợi ý layout: TEXT_ONLY, TEXT_IMAGE, IMAGE_LEFT...


@dataclass
class SlideContent:
    """
    Một slide trong PowerPoint, chứa nhiều block nội dung.
    """
    blocks: List[SlideBlock] = field(default_factory=list[SlideBlock])


@dataclass
class SectionContent:
    """
    Một mục lớn (ví dụ: technical_field, problem, ...) gồm nhiều slide.
    Mỗi slide được mô tả bằng danh sách SlideContent.
    """
    section_id: str                                      # Tên định danh cho mục, ví dụ: "technical_field"
    slides: List[SlideContent] = field(default_factory=list[SlideContent])


@dataclass
class DocumentContent:
    """
    Toàn bộ tài liệu markdown, gồm nhiều mục.
    """
    sections: List[SectionContent] = field(default_factory=list[SectionContent])

if __name__ == "__main__" or os.getenv("UTEST_MODE") == "1":
     document = DocumentContent(
          sections=[
               SectionContent(
                    section_id="technical_field",
                    slides=[
                         SlideContent(
                              blocks=[
                              SlideBlock(
                                   heading="Technical Field",
                                   paragraphs=[
                                        "Botnets using Domain Generation Algorithms (DGA) are capable of generating large numbers of domain names algorithmically.",
                                        "This behavior helps them evade blacklists and traditional static detection systems."
                                   ],
                                   bullet_list=[
                                        "Evades blacklist filters",
                                        "Dynamic domain generation",
                                        "Resilient against domain takedown"
                                   ],
                                   image_path="../_assets/images/dga_overview.png",
                                   layout_hint="TEXT_IMAGE"
                              ),
                              SlideBlock(
                                   heading="Types of DGAs",
                                   paragraphs=[
                                        "DGAs can be categorized into character-based and wordlist-based algorithms."
                                   ],
                                   bullet_list=[
                                        "Character-based DGAs: Random string patterns (e.g., Bamital)",
                                        "Wordlist-based DGAs: Combine dictionary words (e.g., Matsnu)"
                                   ],
                                   layout_hint="TEXT_ONLY"
                              )
                              ]
                         ),
                         SlideContent(
                              blocks=[
                              SlideBlock(
                                   heading="Character-Based DGA",
                                   paragraphs=[
                                        "These DGAs generate domain names using sequences of random characters, resulting in non-human-readable domains."
                                   ],
                                   bullet_list=[
                                        "Hard to predict",
                                        "Difficult to distinguish from legitimate domains"
                                   ],
                                   image_path="../_assets/images/char_dga_example.png",
                                   layout_hint="TEXT_IMAGE"
                              ),
                              SlideBlock(
                                   heading="Real-world Example",
                                   paragraphs=[
                                        "The Bamital botnet used character-based DGA to generate hundreds of domains daily."
                                   ],
                                   layout_hint="TEXT_ONLY"
                              )
                              ]
                         )
                    ]
               )
          ]
     )
