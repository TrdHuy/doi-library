from .base_injector import *
from data.doi_template.v1.contract import *
from data.pptxdata_utils import *

@register_injector
class TitleInjector(BaseSlideInjector):
    inject_id = "title_slide"

    def inject(self, slide: DL_Slide, context: InjectContext):
        basic_info = context.basic_info

        title_shape = self.find_shape_by_name(slide, ShapeName.TITLE)
        if not title_shape or not title_shape.text:
            raise NotImplementedError(
                "Chưa hỗ trợ nếu shape TITLE không có!")

        paragraphs = title_shape.text.paragraphs
        if not paragraphs or not paragraphs[0].runs or not paragraphs[0].runs[0]:
            raise NotImplementedError(
                "Chưa hỗ trợ nếu shape TITLE không có paragraph/run mẫu")

        template_run = paragraphs[0].runs[0]
        new_run = make_run_from_template(template_run, basic_info.title)
        title_shape.text.paragraphs[0].runs[0] = new_run


@register_injector
class BasicInfoInjector(BaseSlideInjector):
    inject_id = "basic_info_slide"

    def inject(self, slide: DL_Slide, context: InjectContext):
        basic_info = context.basic_info

