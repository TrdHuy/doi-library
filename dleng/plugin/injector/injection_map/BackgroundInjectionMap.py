from data.doi_template.v1.slide_data_model import DocumentContent
from data.doi_template.v1.contracts.md_contract import SectionId
from plugin.injector.machine.injector_base import *
from .decorator import inject_with, register_injections
from .InjectionMap import InjectionMap

@register_injections
class BackgroundInjectionMap(InjectionMap):
     def __init__(self, document_content: DocumentContent):
        self.document_content = document_content
        
     def _get_section(self, section_id: SectionId) -> SectionContent:
          for section in self.document_content.sections:
               if section.section_id == section_id.value:
                    return section
          raise ValueError(f"Không tìm thấy SectionContent với section_id = {section_id}")
     
     @inject_with(SectionContentInjector(SlideInjectId.BACKGROUND_TECHNICAL_SLIDE))
     def get_technical_field(self):
          return InjectValue(value=self._get_section(SectionId.TECHNICAL_FIELD))

     @inject_with(SectionContentInjector(SlideInjectId.BACKGROUND_PROBLEM_SLIDE))
     def get_problem_field(self):
          return InjectValue(value=self._get_section(SectionId.PROBLEM))

     @inject_with(SectionContentInjector(SlideInjectId.BACKGROUND_PRIOR_ART_SLIDE))
     def get_prior_art_field(self):
          return InjectValue(value=self._get_section(SectionId.PRIOR_ART))

     @inject_with(SectionContentInjector(SlideInjectId.BACKGROUND_PURPOSE_SLIDE))
     def get_purpose_field(self):
          return InjectValue(value=self._get_section(SectionId.PURPOSE))