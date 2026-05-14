from __future__ import annotations
from abc import ABC
from typing import Any, TypeVar


T = TypeVar('T', bound='VBATypeBase')


class VBATypeBase(ABC):
    value: Any

    def __add__(self: T, other: T) -> T:
        from vba_registry import registry
        return registry.execute_binary_op("+", self, other)
