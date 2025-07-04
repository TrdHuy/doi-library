import json
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

EMU = 1  # đơn vị đã là EMU trong JSON dump

from pptx.enum.dml import MSO_THEME_COLOR

def apply_shape_format(shape_obj, raw_attrs):
    # Fill
    try:
        fill_rgb = raw_attrs.get("fill", {}).get("fore_color")
        if fill_rgb and "RGBColor" in str(fill_rgb):
            r, g, b = map(int, fill_rgb.strip("RGBColor() ").split(","))
            shape_obj.fill.solid()
            shape_obj.fill.fore_color.rgb = RGBColor(r, g, b)
    except: pass

    # Border
    try:
        line_info = raw_attrs.get("line", {})
        line_rgb = line_info.get("color")
        if line_rgb and "RGBColor" in str(line_rgb):
            r, g, b = map(int, line_rgb.strip("RGBColor() ").split(","))
            shape_obj.line.color.rgb = RGBColor(r, g, b)
        width = line_info.get("width")
        if width and isinstance(width, (int, float)):
            shape_obj.line.width = Pt(width)
    except: pass

    # TextFrame
    try:
        tf_attrs = raw_attrs.get("text_frame", {})
        tf = shape_obj.text_frame
        tf.margin_top = tf_attrs.get("margin_top", tf.margin_top)
        tf.margin_bottom = tf_attrs.get("margin_bottom", tf.margin_bottom)
        tf.margin_left = tf_attrs.get("margin_left", tf.margin_left)
        tf.margin_right = tf_attrs.get("margin_right", tf.margin_right)
        tf.auto_size = tf_attrs.get("auto_size", tf.auto_size)
        tf.word_wrap = tf_attrs.get("word_wrap", tf.word_wrap)
    except: pass

def parse_color(json_color_str):
    if not json_color_str or "None" in json_color_str:
        return {"type": "none"}

    if json_color_str.startswith("RGBColor"):
        try:
            r, g, b = map(int, json_color_str.strip("RGBColor() ").split(","))
            return {"type": "rgb", "value": RGBColor(r, g, b)}
        except:
            return {"type": "error"}

    if json_color_str.startswith("Theme:"):
        theme_str = json_color_str.replace("Theme:", "").strip().upper()
        try:
            theme_enum = MSO_THEME_COLOR.__members__.get(theme_str)
            if theme_enum:
                return {"type": "theme", "value": theme_enum}
        except:
            pass

    # Trường hợp hex như "002060"
    try:
        if len(json_color_str) == 6:
            r = int(json_color_str[0:2], 16)
            g = int(json_color_str[2:4], 16)
            b = int(json_color_str[4:6], 16)
            return {"type": "rgb", "value": RGBColor(r, g, b)}
    except:
        pass

    return {"type": "unknown"}

def parse_rgb_color(color_str):
    """Chuyển chuỗi màu (hex hoặc RGBColor) thành RGBColor(r, g, b) hoặc None"""
    if not color_str or "undefined" in color_str.lower():
        return None
    try:
        if color_str.startswith("RGBColor"):
            # RGBColor(0, 0, 0)
            rgb_str = color_str.strip("RGBColor() ")
            r, g, b = map(int, rgb_str.split(","))
            return RGBColor(r, g, b)
        else:
            # Hex dạng "002060"
            color_str = color_str.strip().lstrip("#")
            if len(color_str) == 6:
                r = int(color_str[0:2], 16)
                g = int(color_str[2:4], 16)
                b = int(color_str[4:6], 16)
                return RGBColor(r, g, b)
    except:
        return None

def build_pptx_from_json(json_path, output_pptx):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    prs = Presentation()
    prs.slide_width = data.get("slide_width", prs.slide_width)
    prs.slide_height = data.get("slide_height", prs.slide_height)
    blank_layout = prs.slide_layouts[6]

    for slide_data in data["slides"]:
        slide = prs.slides.add_slide(blank_layout)
        #if slide_data["slide_number"] != 2: # For debug
         #   continue

        for shape_data in slide_data["shapes"]:
            pos = shape_data["position"]
            left, top, width, height = pos["x"], pos["y"], pos["width"], pos["height"]
            shape_type = shape_data.get("type", -1)
            bg_color = parse_rgb_color(shape_data.get("background_fill_color"))

            shape_obj = None

            # TABLE
            if shape_type == 19 and shape_data.get("table") and "data" in shape_data["table"]:
                tbl_data = shape_data["table"]["data"]
                rows, cols = len(tbl_data), len(tbl_data[0])
                shape_obj = slide.shapes.add_table(rows, cols, left, top, width, height)
                tbl = shape_obj.table
                for r in range(rows):
                    for c in range(cols):
                        tbl.cell(r, c).text = tbl_data[r][c]

            # TEXTBOX
            elif shape_data.get("text"):
                shape_obj = slide.shapes.add_textbox(left, top, width, height)
                tf = shape_obj.text_frame
                tf.clear()
                for para_data in shape_data["text"]:
                    p = tf.add_paragraph()
                    for run_data in para_data["runs"]:
                        r = p.add_run()
                        r.text = run_data["text"]

                        font = r.font
                        if run_data.get("font_name"):
                            font.name = run_data["font_name"]
                        if run_data.get("font_size"):
                            font.size = Pt(run_data["font_size"])
                        if run_data.get("bold") is not None:
                            font.bold = run_data["bold"]
                        if run_data.get("italic") is not None:
                            font.italic = run_data["italic"]
                        color_info = parse_color(run_data.get("font_color"))
                        if color_info["type"] == "rgb":
                            font.color.rgb = color_info["value"]
                        elif color_info["type"] == "theme":
                            font.color.theme_color = color_info["value"]

            if not shape_obj:
                continue  # không tạo shape nếu không xử lý được
            
            apply_shape_format(shape_obj, shape_data.get("raw_attributes", {}))
            # Áp dụng màu nền (background_fill_color)
            if bg_color:
                try:
                    shape_obj.fill.solid()
                    shape_obj.fill.fore_color.rgb = bg_color
                except:
                    pass

            # Áp dụng border nếu có
            border = shape_data.get("border", {})
            border_color = parse_rgb_color(border.get("color"))
            if border_color:
                try:
                    shape_obj.line.color.rgb = border_color
                except:
                    pass
            if border.get("width_pt") not in [None, "Default"]:
                try:
                    shape_obj.line.width = Pt(float(border["width_pt"]))
                except:
                    pass

    prs.save(output_pptx)
    print(f"✅ Đã phục hồi PowerPoint: {output_pptx}")

if __name__ == "__main__":
    build_pptx_from_json("bin/dump_output.json", "bin/restored_from_json.pptx")