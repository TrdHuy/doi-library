from .SafeElementWrapper import SafeElementWrapper
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.text.text import _Paragraph  # type: ignore

def get_or_add_pPr(p: _Paragraph) -> SafeElementWrapper:
    """Trả về phần tử pPr XML của paragraph, tạo mới nếu chưa có"""
    p_elem = getattr(p, "_element", None)
    if p_elem is None:
        raise AttributeError("Paragraph does not have an _element attribute")

    pPr = p_elem.find(qn("a:pPr"))
    if pPr is None:
        pPr = OxmlElement("a:pPr")
        p_elem.append(pPr)
    return SafeElementWrapper(pPr)
