from .pptxdata import DL_Text, DL_TextParagraph, DL_Run, DL_TextFrameFormat
from typing import Optional


def make_run_from_values(
    text: str,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_color: Optional[str] = None,
    run_index: int = 0
) -> DL_Run:
    return DL_Run(
        text=text,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        font_color=font_color,
        run_index=run_index
    )

def make_run_from_template(template_run: DL_Run,
                           new_text: str,
                           run_index: int = 0) -> DL_Run:
    return DL_Run(
        text=new_text,
        font_name=template_run.font_name,
        font_size=template_run.font_size,
        bold=template_run.bold,
        italic=template_run.italic,
        font_color=template_run.font_color,
        run_index=run_index
    )