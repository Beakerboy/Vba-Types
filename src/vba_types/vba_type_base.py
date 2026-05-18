from __future__ import annotations
from abc import ABC
from typing import Any, TypeVar


T = TypeVar('T', bound='VBATypeBase')


class VBATypeBase(ABC):
    value: Any

    def __add__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("+", self, other)

    def __sub__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("-", self, other)

    def __mod__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("%", self, other)

    def __mul__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("*", self, other)

    def __pow__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("**", self, other)

    def __truediv__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("/", self, other)

    def __floordiv__(self: T, other: T) -> "VBATypeBase":
        from vba_types.vba_registry import registry
        return registry.execute("//", self, other)
