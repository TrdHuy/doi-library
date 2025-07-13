from enum import Enum

class SectionId(str, Enum):
    TECHNICAL_FIELD = "technical_field"
    PROBLEM = "problem"
    PRIOR_ART = "prior_art"
    PURPOSE = "purpose"