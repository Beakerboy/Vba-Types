from typing import Optional, TypeVar
from vba_types.empty import Empty
from .types_registry import registry


T = TypeVar('T', bound='VBAVariable')


class VBAVariable:
    def __init__(self: T, declared_type: str = "Variant", value: Optional[VBAVariable | VBATypeBase] = None) -> None:
        self._declared_type = declared_type
        if value is None:
            value = Empty
        self.value = value

    def __repr__(self: T):
        return f"VBAVariable({self.declared_type} = {repr(self._value)})"

    def __add__(self: T, other: VBAVariable | VBATypeBase) -> VBATypeBase:
        return self._value + self._unwrap(other)

    def __radd__(self, other):
        return self._value + self._unwrap(other)

    @property
    def declared_type(self: T) -> str:
        return self._declared_type

    @property
    def value(self: T) -> VBATypeBase:
        return self._value

    @value.setter
    def value(self: T, incoming: VBAVariable | VBATypeBase) -> None:
        # Unwrap incoming value if it is another variable container
        if isinstance(incoming, VBAVariable):
            incoming = incoming.value
            
        # Let-coercion logic
        self._value = registry.coerce(incoming, self.declared_type)

     def _unwrap(self: T, other: VBAVariable | VBATypeBase) -> VBATypeBase:
        """Helper to extract the raw VBATypeBase value from a wrapper."""
        if isinstance(other, VBAVariable):
            return other.value
        return other
