from __future__ import annotations
from .vba_type_base import VBATypeBase
from typing import Optional, Type, TypeVar


T = TypeVar("T", bound="VBAEmpty")


class VBAEmpty(VBATypeBase):
    """
    Represents the VBA 'Empty' type.
    It is initialized to 0 in a numeric context and "" in a string context.
    """
    _instance: Optional['VBAEmpty'] = None
    value: int = 0

    def __new__(cls: Type[T]) -> 'VBAEmpty':
        if cls._instance is None:
            cls._instance = super(VBAEmpty, cls).__new__(cls)
        return cls._instance

    def __repr__(self: T) -> str:
        return "Empty"

    def __str__(self: T) -> str:
        return ""

    def __int__(self: T) -> int:
        return 0

    def __float__(self: T) -> float:
        return 0.0

    def __bool__(self: T) -> bool:
        # In VBA, Empty evaluates to False/0
        return False


Empty = VBAEmpty()
