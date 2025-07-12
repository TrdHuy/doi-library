import json
import os
import re
from markdown import markdown
from bs4 import BeautifulSoup
import pandas as pd
from data.doi_template.v1.basic_info import *
from data.doi_template.v1.background import *
from io import StringIO


def is_markdown_table(text: str) -> bool:
    return re.match(r"^\s*\|.*\|\s*\n\s*\|[-| ]+\|", text.strip(), re.MULTILINE) is not None


def parse_table(text: str, section_name: str = "unknown") -> pd.DataFrame:
    try:
        lines = text.strip().splitlines()
        if len(lines) < 3:
            raise ValueError("Bảng markdown quá ngắn (phải ≥ 3 dòng)")

        header_line = lines[0]
        data_lines = lines[2:]  # Bỏ dòng separator
        data_csv = "\n".join([header_line] + data_lines)

        df = pd.read_csv(
            StringIO(data_csv),
            sep="|",
            engine="python",
            skipinitialspace=True,
            dtype=str,  # Đảm bảo mọi thứ là string để không lỗi chuyển kiểu
            on_bad_lines="error"
        ).dropna(axis=1, how="all").dropna(axis=0, how="all")

        df.columns = [col.strip() for col in df.columns]
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

        return df.reset_index(drop=True)

    except Exception as e:
        print(f"\n❌ Lỗi khi parse bảng trong section: `{section_name}`")
        print("📄 Nội dung bảng:")
        print("-" * 40)
        print(text)
        print("-" * 40)
        print("🔍 Exception:", e)
        raise ValueError(
            f"Parse thất bại ở bảng section `{section_name}`: {e}")


def extract_basic_info_sections(basic_info_path_md: str) -> Optional[BasicInfo]:
    with open(basic_info_path_md, encoding="utf-8") as f:
        markdown_text = f.read()
    pattern = r"^### .*?$"  # match section headers
    lines = markdown_text.strip().splitlines()
    sections = {}
    current_section = None
    buffer = []

    for line in lines:
        if re.match(pattern, line):
            if current_section:
                body = "\n".join(buffer).strip()
                if is_markdown_table(body):
                    table = parse_table(body, section_name=current_section)
                    sections[current_section] = table if table is not None else body
                else:
                    sections[current_section] = body
                buffer = []
            current_section = line.lstrip("#").strip().split(" ", 1)[-1]
        elif line.strip() == "---":
            continue
        else:
            buffer.append(line)

    # Lưu section cuối cùng
    if current_section:
        body = "\n".join(buffer).strip()
        if is_markdown_table(body):
            table = parse_table(body, section_name=current_section)
            sections[current_section] = table if table is not None else body
        else:
            sections[current_section] = body

    return parse_basic_info(sections)


# with open(r"C:\Users\Hp\Desktop\temp\python\doi-library\doi_src\template_project_doi\0_basic_info\basic_info.md", encoding="utf-8") as f:
#     md = f.read()

# data = extract_sections(md)
# basic = parse_basic_info(data)

def extract_config_from_line(line: str) -> (str, Optional[dict]):
    config_match = re.search(r'<!--\s*(\{.*?\})\s*-->', line)
    if config_match:
        config_str = config_match.group(1)
        try:
            config = json.loads(config_str)
            clean_line = re.sub(r'<!--\s*\{.*?\}\s*-->', '', line)
            return clean_line, config
        except json.JSONDecodeError:
            pass
    return line, None


def parse_markdown_to_blocks(md_text: str) -> List[ContentBlock]:
    blocks = []
    lines = md_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line:
            i += 1
            continue
        
        # Kiểm tra nếu là bảng
        if '|' in line and re.match(r'^\|.*\|$', line):
            table_lines = [line]
            i += 1
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i].strip())
                i += 1

            # Parse bảng
            table_data = []
            for row_line in table_lines:
                row = [cell.strip() for cell in row_line.strip('|').split('|')]
                table_data.append(row)
            # Bỏ dòng header line (vd: ---|---|---)
            if len(table_data) >= 2 and re.match(r'^-+$', table_data[1][0]):
                table_data.pop(1)

            blocks.append(ContentBlock(
                type="table",
                table_data=table_data
            ))
            continue
        
        i += 1
        line, config = extract_config_from_line(line)
        # Hình ảnh
        img_match = re.match(r'!\[(.*)\]\((.*)\)', line)
        if img_match:
            blocks.append(ContentBlock(
                type="image", image_path=img_match.group(2),
                config=config))
            continue

        bullet_match = re.match(r'^(\s*)([-*+])\s+((\d+):\s+)?(.*)', line)
        if bullet_match:
            indent = bullet_match.group(1)
            level = len(indent) // 2
            bullet_char = bullet_match.group(4)  # None nếu không có
            text = bullet_match.group(5)
            blocks.append(ContentBlock(
                type="bullet",
                text=text.strip(),
                level=level,
                bullet_char=bullet_char,  # "1", "2",... hoặc None
                config=config
            ))
            continue

        # Heading (title phụ)
        heading_match = re.match(r'^(#{2,6})\s+(.*)', line)
        if heading_match:
            blocks.append(ContentBlock(type="title",
                                       text=heading_match.group(2),
                                       config=config))
            continue

        # Heading (title chính)
        heading_match = re.match(r'^(#)\s+(.*)', line)
        if heading_match:
            continue

        # Đoạn văn thường
        blocks.append(ContentBlock(type="paragraph",
                                   text=line,
                                   config=config))
    return blocks


def load_background_from_folder(folder_path: str) -> Optional[Background]:
    sections = []

    for section_name in os.listdir(folder_path):
        section_path = os.path.join(folder_path, section_name)
        slide_items = []
        if os.path.isdir(section_path):
            # Trường hợp chia slide nhỏ (vd: technical_field/dga_botnet.md)
            md_files = [f for f in os.listdir(
                section_path) if f.endswith('.md')]
            if md_files:
                for md_file in md_files:
                    slide_id = md_file.replace(".md", "")
                    md_path = os.path.join(section_path, md_file)
                    json_path = os.path.join(section_path, slide_id + ".json")

                    with open(md_path, "r", encoding="utf-8") as f:
                        content = f.read()
                    blocks = parse_markdown_to_blocks(content)

                    config = None
                    if os.path.exists(json_path):
                        with open(json_path, "r", encoding="utf-8") as f:
                            config = json.load(f)

                    slide_items.append(BackgroundSlide(
                        id=slide_id,
                        title=None,  # hoặc lấy từ block đầu tiên nếu là title
                        blocks=blocks,
                        config=config
                    ))

        # Trường hợp chỉ có 1 file .md + .json (vd: problem.md chứa nhiều slide)
        else:
            section_name = os.path.splitext(os.path.basename(section_path))[0]
            base_md = os.path.join(folder_path, section_name + ".md")
            base_json = os.path.join(folder_path, section_name + ".json")
            if os.path.exists(base_md):
                with open(base_md, "r", encoding="utf-8") as f:
                    raw_md = f.read()
                slide_chunks = raw_md.split("</br>")
                config_data = []
                if os.path.exists(base_json):
                    with open(base_json, "r", encoding="utf-8") as f:
                        config_data = json.load(f)

                for i, chunk in enumerate(slide_chunks):
                    slide_id = f"{section_name}_slide_{i+1}"
                    blocks = parse_markdown_to_blocks(chunk.strip())
                    cfg = config_data[i] if i < len(config_data) else None
                    slide_items.append(BackgroundSlide(
                        id=slide_id,
                        blocks=blocks,
                        config=cfg
                    ))

        sections.append(BackgroundSection(
            name=section_name,
            slides=slide_items
        ))

    return Background(sections=sections)


# bg = load_background_from_folder(
#     r"C:\Users\huy.td1\Desktop\Temp\doi-library\doi_src\template_project_doi\1_background")
# bg
