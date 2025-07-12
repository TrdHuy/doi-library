from pptx.slide import Slide
from pptx.shapes.autoshape import Shape
from pptx.dml.color import RGBColor

def get_slide_visible(slide: Slide) -> bool:
    """
    Kiểm tra xem slide có đang được hiển thị hay bị ẩn.
    (ẩn bằng tính năng Hide Slide trong PowerPoint)
    """
    try:
        sp = slide._element
        cSld = sp.find("p:cSld", namespaces=sp.nsmap)
        if cSld is not None:
            show_attr = cSld.get("show")
            return show_attr != "0"  # "0" nghĩa là ẩn
    except:
        pass
    return True  # mặc định là hiển thị

def set_slide_visible(slide: Slide, visible: bool):
    """
    Đặt trạng thái visible (ẩn/hiện) cho slide.
    """
    sp = slide._element
    cSld = sp.find("p:cSld", namespaces=sp.nsmap)
    if cSld is not None:
        if not visible:
            cSld.set("show", "0")  # ẩn slide
        elif "show" in cSld.attrib:
            del cSld.attrib["show"]  # xoá để trở về mặc định là hiển thị

def get_shape_visible(shape) -> bool:
    """
    Trả về True nếu shape đang hiển thị, False nếu bị ẩn trong Selection Pane (cNvPr@hidden="1")
    """
    try:
        cNvPr = shape._element.find(".//p:cNvPr", namespaces={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
        if cNvPr is not None:
            return cNvPr.get("hidden") != "1"
        return True
    except Exception as e:
        print(f"[get_shape_visible] ⚠️ Lỗi khi lấy trạng thái visible: {e}")
        return True
    
def set_shape_visible(shape, visible: bool):
    """
    Đặt trạng thái hiển thị của shape (ẩn/hiện trong Selection Pane).
    """
    try:
        sp = shape._element
        cNvPr = sp.find(".//p:cNvPr", namespaces=sp.nsmap)
        if cNvPr is not None:
            if visible:
                if "hidden" in cNvPr.attrib:
                    del cNvPr.attrib["hidden"]
            else:
                cNvPr.set("hidden", "1")
    except Exception as e:
        print(f"[set_shape_visible] ⚠️ Lỗi khi set trạng thái visible: {e}")
        
def lock_shape_position(shape: Shape):
    """
    Khóa không cho shape di chuyển hoặc resize bằng cách thêm <a:spLocks noMove="1" noResize="1"/>
    """
    try:
        spPr = shape._element.find(".//p:spPr", namespaces={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
        if spPr is None:
            print("⚠️ Không tìm thấy p:spPr trong shape.")
            return

        # Thêm thẻ <a:spLocks noMove="1" noResize="1"/>
        from lxml import etree
        NSMAP = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        spLocks = etree.SubElement(spPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}spLocks")
        spLocks.set("noMove", "1")
        spLocks.set("noResize", "1")

        print("✅ Đã khóa shape thành công.")
    except Exception as e:
        print(f"[lock_shape_position] ❌ Lỗi: {e}")
        
def unlock_shape_position(shape: Shape):
    try:
        spPr = shape._element.find(".//p:spPr", namespaces={"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
        if spPr is None:
            return

        for spLocks in spPr.findall(".//a:spLocks", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}):
            spPr.remove(spLocks)
    except Exception as e:
        print(f"[unlock_shape_position] ❌ Lỗi: {e}")

def extract_font_size_from_empty_paragraph(paragraph):
    p_element = paragraph._p
    end_para_rpr = p_element.find(".//a:endParaRPr", namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'})
    if end_para_rpr is not None and "sz" in end_para_rpr.attrib:
        sz = int(end_para_rpr.attrib["sz"])
        return sz / 100  # chuyển thành pt
    return None

def extract_font_info_from_end_para(paragraph):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    p_element = paragraph._p
    rpr = p_element.find(".//a:endParaRPr", namespaces=ns)
    if rpr is None:
        return None

    # Font size
    sz = rpr.get("sz")
    font_size = int(sz) / 100 if sz else None

    # Bold and italic
    bold = rpr.get("b") == "1"
    italic = rpr.get("i") == "1"

    # Font color
    color_node = rpr.find(".//a:srgbClr", namespaces=ns)
    font_color = f"#{color_node.get('val')}" if color_node is not None else None

    # Font name
    latin_node = rpr.find(".//a:latin", namespaces=ns)
    font_name = latin_node.get("typeface") if latin_node is not None else None

    return {
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "font_color": font_color,
        "font_name": font_name,
    }