from .pptxdata import DL_Text, DL_TextParagraph, DL_Run, DL_TextFrameFormat, DL_Table
from typing import Tuple, Optional


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

def set_table_cell_text(
    table: DL_Table,
    row: int,
    col: int,
    text_content: str,
):
    # 1. Kiểm tra hợp lệ
    if row >= table.rows or col >= table.cols:
        raise IndexError(f"Cell ({row}, {col}) is out of table bounds ({table.rows}x{table.cols})")

    # 2. Ghi text vào phần data (text thuần)
    table.data[row][col] = text_content

    text_detail = table.data_detail[row][col]
    if text_detail is None:
        raise IndexError(f"Cell ({row}, {col}) không có text detail, có thể đây là cell merge!")

    # 3. Làm việc với paragraph đầu tiên
    if not text_detail.paragraphs:
        raise ValueError(f"Cell ({row}, {col}) không có paragraph nào trong text detail.")

    paragraph = text_detail.paragraphs[0]

    # 4. Tạo hoặc cập nhật run
    if paragraph.runs:
        # Cập nhật run đầu tiên
        paragraph.runs[0].text = text_content
        paragraph.runs[0].run_index = 1
        # Xoá các run còn lại
        paragraph.runs = [paragraph.runs[0]]
    else:
        # Tạo run mới dựa trên style của paragraph
        run = DL_Run(
            text=text_content,
            font_name=paragraph.font_name,
            font_size=paragraph.font_size,
            bold=paragraph.bold,
            italic=paragraph.italic,
            font_color=paragraph.font_color,
            run_index=1
        )
        paragraph.runs = [run]

    # 5. Xoá các paragraph còn lại nếu có
    text_detail.paragraphs = [paragraph]

def find_table_cell_by_text(table: DL_Table, target_text: str) -> Optional[Tuple[int, int]]:
    """
    Tìm vị trí (row, col) trong bảng DL_Table chứa đúng chuỗi `target_text`.

    Args:
        table (DL_Table): bảng dữ liệu cần tìm
        target_text (str): chuỗi cần tìm

    Returns:
        Tuple[int, int]: vị trí của cell chứa text, hoặc None nếu không tìm thấy
    """
    for row_idx in range(table.rows):
        for col_idx in range(table.cols):
            cell_text = table.data[row_idx][col_idx]
            if cell_text.strip() == target_text.strip():
                return (row_idx, col_idx)
    return None
   