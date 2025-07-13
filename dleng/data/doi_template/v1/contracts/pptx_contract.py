from enum import Enum

class SlideInjectId(str, Enum):
    TITLE_SLIDE = "title_slide"
    BASIC_INFO_SLIDE = "basic_info_slide"
    
    BACKGROUND_TECHNICAL_SLIDE = "background_technical_slide"
    BACKGROUND_PROBLEM_SLIDE = "background_problem_slide"
    BACKGROUND_PRIOR_ART_SLIDE = "background_prior_art_slide"
    BACKGROUND_PURPOSE_SLIDE = "background_purpose_slide"

class ShapeName(str, Enum):
    TITLE                   = "ELE_TITLE_SHAPE"
    BASIC_INFO_TABLE        = "ELE_BASICINFO_TABLE"
    
    PARAGRAPH_CONTENT_AREA   = "ELE_PARAGRAPH_CONTENT_AREA"
    IMAGE_CONTENT_AREA       = "ELE_IMAGE_CONTENT_AREA"

class RunSample(str, Enum):
    DEPARTMENT_RS       = "ELE_DEPARTMENT_RUN_SAMPLE"
    PROJECT_NAME_RS     = "ELE_PROJECTNAME_RUN_SAMPLE"
    INVENTION_TITLE_RS  = "ELE_INVENTION_TITLE_RUN_SAMPLE"
    INVENTOR_NAME_RS    = "ELE_INV_NAME_RUN_SAMPLE"
    CONTRIBUTE_RATE_RS  = "ELE_CR_RUN_SAMPLE"
    EMPLOYEE_ID_RS      = "ELE_EM_ID_RUN_SAMPLE"
    EMPLOYEE_STAT_RS    = "ELE_EM_STAT_RUN_SAMPLE"
