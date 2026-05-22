from __future__ import annotations
from abc import ABC
from typing import Any, TypeVar


T = TypeVar('T', bound='VBATypeBase')


class VBATypeBase(ABC):
    value: Any

    def __add__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("+", self, other)

    def __sub__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("-", self, other)

    def __mod__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("%", self, other)

    def __mul__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("*", self, other)

    def __pow__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("**", self, other)

    def __truediv__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("/", self, other)

    def __floordiv__(self: T, other: T) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("//", self, other)

    def __eq__(self: T,
               other: object) -> bool | "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("==", self, other)

    def __ne__(self: T,                                # type: ignore[override]
               other: object) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("<>", self, other)

    def __lt__(self: T, other: object) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("<", self, other)

    def __gt__(self: T, other: object) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute(">", self, other)

    def __ge__(self: T, other: object) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("=>", self, other)

    def __le__(self: T, other: object) -> "VBATypeBase":
        if not isinstance(other, VBATypeBase):
            return NotImplemented
        from vba_types.vba_registry import registry
        return registry.execute("<=", self, other)
