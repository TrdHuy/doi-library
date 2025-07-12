from pptx.oxml.xmlchemy import BaseOxmlElement
from typing import Any


class SafeElementWrapper:
    def __init__(self, element: BaseOxmlElement):
        self._element: BaseOxmlElement = element

    def set(self, key: str, value: str) -> None:
        self._element.set(key, value)  # type: ignore

    def append(self, child: "SafeElementWrapper") -> None:
        self._element.append(child.get())  # type: ignore

    def get(self) -> BaseOxmlElement:
        return self._element

    def __getattr__(self, name: str) -> Any:
        return getattr(self._element, name)
