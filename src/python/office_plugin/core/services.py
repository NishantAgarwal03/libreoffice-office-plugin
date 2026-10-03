"""Tiny service container: features ask for shared helpers by name."""


class ServiceContainer:
    def __init__(self):
        self._factories = {}
        self._instances = {}

    def register(self, name, factory):
        """`factory` is a zero-arg callable; built lazily on first get()."""
        self._factories[name] = factory

    def register_instance(self, name, instance):
        self._instances[name] = instance

    def has(self, name):
        return name in self._instances or name in self._factories

    def get(self, name):
        if name not in self._instances:
            if name not in self._factories:
                raise KeyError("unknown service %r" % name)
            self._instances[name] = self._factories[name]()
        return self._instances[name]
