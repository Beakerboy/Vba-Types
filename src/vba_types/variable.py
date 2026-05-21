from __future__ import annotations
from typing import Optional, TypeGuard, TypeVar
from vba_types.empty import Empty
from .types_registry import registry
from .vba_type_base import VBATypeBase
from .boolean import VBABoolean
from .numeric_type import VBANumericType
from .string import VBAString


T = TypeVar('T', bound='VBAVariable')


class VBAVariable:
    def __init__(self: T,
                 declared_type: str = "variant",
                 value: Optional[T | VBATypeBase] = None) -> None:
        self._declared_type = declared_type.lower()
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

    def __eq__(self: T,                                # type: ignore[override]
               other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(False)
        if self._is_number_and_string(other):
            if self._declared_type == "variant":
                try:
                    registry.coerce(other._declared_type, self._value) == other
                except Exception:
                    self == registry.coerce(self._declared_type, other._value)
            else:
                try:
                    self == registry.coerce(self._declared_type, other._value)
                except Exception:
                    registry.coerce(other._declared_type, self._value) == other
        return self._value == self._unwrap(other)

    def __ne__(self: T,                                # type: ignore[override]
               other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(True)
        return self._value == self._unwrap(other)

    def __gt__(self: T, other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(isinstance(self._value, VBAString))
        return self._value > self._unwrap(other)

    def __lt__(self: T, other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(issubclass(type(self._value), VBANumericType))
        if self._is_number_and_string(other):
            if self._declared_type == "variant":
                try:
                    coerce = registry.coerce(other._declared_type, self._value)
                    result = coerce < other
                except Exception:
                    result = self < registry.coerce("integer", other._value)
            else:
                if str(other._value) == "":
                    return VBABoolean(True)
                try:
                    coerce = registry.coerce(self._declared_type, other._value)
                    result = self < coerce
                except Exception:
                    result = registry.coerce("integer", self._value) < other
            return VBABoolean(result)
        return self._value < self._unwrap(other)

    def __ge__(self: T, other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(isinstance(self._value, VBAString))
        return self._value >= self._unwrap(other)

    def __le__(self: T, other: object) -> VBATypeBase:
        if not self._is_vba_type(other):
            return NotImplemented
        if self._string_numeric_case(other):
            return VBABoolean(issubclass(type(self._value), VBANumericType))
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
        if self._declared_type == "variant":
            self._value = incoming
        else:
            # Let-coercion logic
            self._value = registry.coerce(self.declared_type, incoming)

    def _unwrap(self: T, other: 'VBAVariable' | VBATypeBase) -> VBATypeBase:
        """Helper to extract the raw VBATypeBase value from a wrapper."""
        if isinstance(other, VBAVariable):
            return other.value
        return other

    def _is_number_and_string(self: T, other: object) -> bool:
        return (
            isinstance(other, VBAVariable) and
            (
                issubclass(type(self._value), VBANumericType) or
                issubclass(type(other._value), VBANumericType)
            ) and
            (
                isinstance(self._value, VBAString) or
                isinstance(other._value, VBAString)
            )
        )

    def _string_numeric_case(self: T, other: object) -> bool:
        """
        If both are variant, and one argument is numeric, and one is a
        string, the number is always smaller.
        """
        return (
            self._declared_type == "variant" and
            isinstance(other, VBAVariable) and
            other._declared_type == "variant" and
            (
                issubclass(type(self._value), VBANumericType) or
                issubclass(type(other._value), VBANumericType)
            ) and
            (
                isinstance(self._value, VBAString) or
                isinstance(other._value, VBAString)
            )
        )

    def _is_vba_type(self: T,
                     other: object) -> TypeGuard['VBAVariable' | VBATypeBase]:
        return isinstance(other, VBAVariable) or isinstance(other, VBATypeBase)
