inject_registry = []

def inject_with(injector):
    def wrapper(func):
        inject_registry.append((func, injector))
        return func
    return wrapper