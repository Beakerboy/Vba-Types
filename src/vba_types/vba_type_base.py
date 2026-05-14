from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Any, TypeVar
from enum import Enum, auto


class VBAType(Enum):
    ARRAY = auto()
    BOOLEAN = auto()
    BYTE = auto()
    CURRENCY = auto()
    DATE = auto()
    DOUBLE = auto()
    EMPTY = auto()
    INTEGER = auto()
    LONG = auto()
    LONGLONG = auto()
    NULL = auto()
    STRING = auto()


T = TypeVar('T', bound='VBATypeBase')


class VBATypeBase(ABC):
    value: Any

    # In your base class:
    @property
    @abstractmethod
    def type_name(self) -> VBAType:
        pass

    def __add__(self: T, other: T) -> T:
        from vba_types.vba_registry import registry
        return registry.execute_binary_op("+", self, other)
