from __future__ import annotations
from typing import Optional, TypeVar
from vba_types.empty import Empty
from .types_registry import registry
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBAVariable')


class VBAVariable:
    def __init__(self: T,
                 declared_type: str = "Variant",
                 value: Optional[T | VBATypeBase] = None) -> None:
        self._declared_type = declared_type
        if value is None:
            value = Empty
        self.value = value

    def __repr__(self: T) -> str:
        return f"VBAVariable({self.declared_type} = {repr(self._value)})"

    def __add__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value + self._unwrap(other)

    def __radd__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value + self._unwrap(other)

    def __sub__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value - self._unwrap(other)

    def __rsub__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._unwrap(other) - self._value

    def __mul__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value * self._unwrap(other)

    def __rmul__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value * self._unwrap(other)

    def __pow__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value ** self._unwrap(other)

    def __rpow__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._unwrap(other) ** self._value

    def __eq__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value == self._unwrap(other)

    def __gt__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value > self._unwrap(other)

    def __lt__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value < self._unwrap(other)

    def __le__(self: T, other: T | VBATypeBase) -> VBATypeBase:
        return self._value <= self._unwrap(other)

    @property
    def declared_type(self: T) -> str:
        return self._declared_type

    @property
    def value(self: T) -> VBATypeBase:
        return self._value

    @value.setter
    def value(self: T, incoming: T | VBATypeBase) -> None:
        # Unwrap incoming value if it is another variable container
        if isinstance(incoming, VBAVariable):
            incoming = incoming.value
        if self._declared_type == "Variant":
            self._value = incoming
        else:
            # Let-coercion logic
            self._value = registry.coerce(self.declared_type, incoming)

    def _unwrap(self: T, other: T | VBATypeBase) -> VBATypeBase:
        """Helper to extract the raw VBATypeBase value from a wrapper."""
        if isinstance(other, VBAVariable):
            return other.value
        return other
