from enum import Enum
from dataclasses import dataclass, field
from typing import TypeVar, Generic, Optional, Any

T = TypeVar('T')

class InjectMetaKey(str, Enum):
    INSERT_INDEX = "INSERT_INDEX"
    TEMPLATE_ROW_INDEX = "TEMPLATE_ROW_INDEX"
    IS_DELETE_TEMPLATE_ROW = "IS_DELETE_TEMPLATE_ROW"


@dataclass
class InjectValue(Generic[T]):
    value: T
    __meta: dict[str, Any] = field(
        default_factory=dict[str, Any], init=False, repr=False)

    def __init__(self, value: Any, meta: Optional[dict[InjectMetaKey, Any]] = None):
        self.value = value
        self.__meta = {k.value: v for k, v in meta.items()} if meta else {}

    def __getitem__(self, key: InjectMetaKey) -> Optional[Any]:
        return self.__meta.get(key.value)

    def __setitem__(self, key: InjectMetaKey, val: Any) -> None:
        self.__meta[key.value] = val

    def __contains__(self, key: InjectMetaKey) -> bool:
        return key.value in self.__meta

    def remove(self, key: InjectMetaKey) -> None:
        self.__meta.pop(key.value, None)

    def keys(self):
        return [InjectMetaKey(k) for k in self.__meta]

    def get(self, key: InjectMetaKey, default: Any = None) -> Any:
        return self.__meta.get(key.value, default)

    def get_int(self, key: InjectMetaKey, default: int = 0) -> int:
        val = self.get(key, default)
        try:
            return int(val)
        except (ValueError, TypeError):
            return default

    def __repr__(self):
        return f"InjectValue(value={self.value})"
