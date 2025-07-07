import re
from markdown import markdown
from bs4 import BeautifulSoup
import pandas as pd
from data.doi_template.v1.basic_info import *
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
        raise ValueError(f"Parse thất bại ở bảng section `{section_name}`: {e}")

def extract_sections(markdown_text: str):
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

    return sections


# with open(r"C:\Users\Hp\Desktop\temp\python\doi-library\doi_src\template_project_doi\0_basic_info\basic_info.md", encoding="utf-8") as f:
#     md = f.read()

# data = extract_sections(md)
# basic = parse_basic_info(data)
