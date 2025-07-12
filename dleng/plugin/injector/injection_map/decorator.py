from plugin.injector.machine.injector_base import INJECT_REGISTRY
from typing import Callable, Any

def inject_with(injector: Any) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    def wrapper(func: Callable[..., Any]) -> Callable[..., Any]:
        INJECT_REGISTRY.append((func, injector))
        return func
    return wrapper