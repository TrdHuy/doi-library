
from data.doi_template.v1.contracts.pptx_contract import ShapeName, SlideInjectId
from data.doi_template.v1.slide_data_model import SectionContent
from data.pptxdata import DL_PPTXData
from plugin.injector.machine.inject_value import InjectValue
from plugin.injector.machine.injector_base import Injector, TemplateShapeTextInjector

class SectionContentInjector(Injector[SectionContent]):
    def __init__(self, slide_id: SlideInjectId,
                 text_content_shape_name: ShapeName,
                 image_content_shape_name: ShapeName):
        self.slide_id = slide_id
        self.text_content_shape_name = text_content_shape_name
        self.image_content_shape_name = image_content_shape_name
        
    def inject(self, pptx_data: DL_PPTXData, injected_value: InjectValue[SectionContent]) -> None:
        section: SectionContent = injected_value.value

        for slide_index, slide_content in enumerate(section.slides):
            if slide_index >= len(slide_content.blocks):
                continue

            for block in slide_content.blocks:
                block_value = InjectValue[str](value="s")

                injector = TemplateShapeTextInjector(
                    slide_id=self.slide_id,
                    shape_name=self.text_content_shape_name
                )
                injector.inject(pptx_data, block_value)

    def build_dl_text(self, block: SlideBlock) -> DL_Text:
        lines = []

        if block.heading:
            lines.append(DL_TextLine(text=block.heading, is_bold=True))

        for item in block.items:
            lines.append(
                DL_TextLine(
                    text=item.text,
                    is_bullet=(item.type == "bullet"),
                    bullet_level=item.level if item.type == "bullet" else 0
                )
            )

        paragraph = DL_TextParagraph(lines=lines)
        return DL_Text(paragraphs=[paragraph])