from data.pptxdata import *
from data.doi_template.v1.basic_info import *
from data.doi_template.v1.background import *
INJECTOR_REGISTRY = {}


def register_injector(cls):
    INJECTOR_REGISTRY[cls.inject_id] = cls()
    return cls


class InjectContext:
    def __init__(self,
                 basic_info: BasicInfo = None,
                 background: Background = None,
                 prior_art=None):
        self.basic_info = basic_info
        self.background = background
        self.prior_art = prior_art


class BaseSlideInjector:
    inject_id: str

    def inject(self, slide: DL_Slide, data_context: InjectContext):
        raise NotImplementedError

    def find_shape_by_name(self, slide: DL_Slide, shape_name: str) -> Optional[DL_Shape]:
        """
        Trả về shape có tên khớp với shape_name (so sánh với shape.shape_name).
        """
        return next((s for s in slide.shapes if s.shape_name == shape_name), None)
