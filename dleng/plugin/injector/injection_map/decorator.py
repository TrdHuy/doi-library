from plugin.injector.machine.injector_base import INJECT_REGISTRY, Injector, InjectValue, T
from typing import Callable, Any, Type, cast
from .InjectionMap import InjectionMap

def register_injections(cls: Type[InjectionMap]) -> Type[InjectionMap]:
    """
    Tự động gom toàn bộ method có attribute __injector__ trong class
    và gán vào INJECT_REGISTRY tương ứng class đó.
    """
    for attr_name in dir(cls):
        attr = getattr(cls, attr_name)
        if callable(attr) and hasattr(attr, "__injector__"):
            method = cast(Callable[..., Any], attr)
            injector = getattr(method, "__injector__")
            INJECT_REGISTRY.setdefault(cls, []).append((method, injector))
    return cls


def inject_with(injector: Injector[T]) -> Callable[[Callable[..., InjectValue[T]]], Callable[..., InjectValue[T]]]:
    def wrapper(func: Callable[..., InjectValue[T]]) -> Callable[..., InjectValue[T]]:
        setattr(func, "__injector__", injector)
        return func
    return wrapper
