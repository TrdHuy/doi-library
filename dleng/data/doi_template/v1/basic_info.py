from dataclasses import dataclass
from typing import List
import pandas as pd

@dataclass
class Inventor:
    no: int
    full_name: str
    contribution: str  # % string
    employee_no: str
    status: str

@dataclass
class Abbreviation:
    abbreviation: str
    description: str

@dataclass
class AppendixItem:
    name: str
    description: str

@dataclass
class Disclosure:
    technical_field: str
    problem: str
    prior_art: str
    purpose: str

@dataclass
class BasicInfo:
    title: str
    invention_number: str
    date_received: str
    review_decision: str
    department: str
    project_name: str
    invention_title: str
    inventors: List[Inventor]
    disclosure: Disclosure
    abbreviations: List[Abbreviation]
    appendix: List[AppendixItem]

def parse_basic_info(sections: dict) -> BasicInfo:
    # 1. Title
    title = sections.get("title")
    if not title:
        raise ValueError("Thiếu mục `title`")

    # 2. Invention Number
    patent_text = sections.get("for_patent_team_use", "")
    lines = patent_text.splitlines()
    inv_number = next((l.split(":", 1)[-1].strip() for l in lines if "Invention Number" in l), None)
    date_recv = next((l.split(":", 1)[-1].strip() for l in lines if "Date Received" in l), None)
    review = next((l.split(":", 1)[-1].strip() for l in lines if "Review Decision" in l), None)
#     if not (inv_number and date_recv and review):
#         raise ValueError("Sai định dạng ở `for_patent_team_use`")

    # 3. Các trường đơn
    department = sections.get("department")
    project_name = sections.get("project_name")
    invention_title = sections.get("invention_title")
    if not (department and project_name and invention_title):
        raise ValueError("Thiếu trường `department`, `project_name` hoặc `invention_title`")

    # 4. Inventors
    df = sections.get("inventors")
    if not isinstance(df, pd.DataFrame):
        raise ValueError("`inventors` không phải bảng")
    expected_cols = ["No.", "Full Name", "Contribution", "Employee No.", "Status"]
    for col in expected_cols:
        if col not in df.columns:
            raise ValueError(f"Bảng inventors thiếu cột '{col}'")
    inventors = [
        Inventor(
            no=int(row["No."]),
            full_name=row["Full Name"],
            contribution=row["Contribution"],
            employee_no=row["Employee No."],
            status=row["Status"]
        ) for _, row in df.iterrows()
    ]

    # 5. Disclosure
    disclosure_raw = sections.get("disclosure_of_invention", "")
    disclosure_lines = [l.strip("- ").strip() for l in disclosure_raw.splitlines() if l.strip()]
    disclosure_map = {}
    for l in disclosure_lines:
        if ":" in l:
            k, v = l.split(":", 1)
            disclosure_map[k.strip().lower()] = v.strip()
    for field in ["technical field", "problem", "prior art", "purpose"]:
        if field not in disclosure_map:
            raise ValueError(f"Disclosure thiếu trường `{field}`")
    disclosure = Disclosure(
        technical_field=disclosure_map["technical field"],
        problem=disclosure_map["problem"],
        prior_art=disclosure_map["prior art"],
        purpose=disclosure_map["purpose"]
    )

    # 6. Abbreviations
    df_abbr = sections.get("abbreviation_table")
    if not isinstance(df_abbr, pd.DataFrame):
        raise ValueError("`abbreviation_table` không phải bảng")
    abbreviations = [
        Abbreviation(abbreviation=row[0], description=row[1])
        for _, row in df_abbr.iterrows()
    ]

    # 7. Appendix
    df_app = sections.get("appendix_table")
    if not isinstance(df_app, pd.DataFrame):
        raise ValueError("`appendix_table` không phải bảng")
    appendix = [
        AppendixItem(name=row[0], description=row[1])
        for _, row in df_app.iterrows()
    ]

    # ✅ Return kết quả
    return BasicInfo(
        title=title,
        invention_number=inv_number,
        date_received=date_recv,
        review_decision=review,
        department=department,
        project_name=project_name,
        invention_title=invention_title,
        inventors=inventors,
        disclosure=disclosure,
        abbreviations=abbreviations,
        appendix=appendix
    )
