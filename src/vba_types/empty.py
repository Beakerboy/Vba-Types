from typing import Any, Optional, TypeVar, Union
from .exceptions import DivisionByZeroError
from .vba_type_base import VBATypeBase


T = TypeVar("T", bound="VBAEmpty")


class VBAEmpty(VBATypeBase):
    """
    Represents the VBA 'Empty' type. 
    It is initialized to 0 in a numeric context and "" in a string context.
    """
    _instance: Optional['VBAEmpty'] = None

    def __new__(cls) -> 'VBAEmpty':
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

    def __eq__(self: T, other: Any) -> bool:
        if isinstance(other, VBAEmpty):
            return True
        if isinstance(other, (int, float)):
            return other == 0
        if isinstance(other, str):
            return other == ""
        return False

    # Arithmetic behavior (Empty acts as 0)
    def __add__(self: T, other: Any) -> Any: return 0 + other
    def __radd__(self: T, other: Any) -> Any: return other + 0
    def __sub__(self: T, other: Any) -> Any: return 0 - other
    def __rsub__(self: T, other: Any) -> Any: return other - 0
    def __mul__(self: T, other: Any) -> Any: return 0 * other
    def __rmul__(self: T, other: Any) -> Any: return other * 0
    def __truediv__(self: T, other: Any) -> Any: return 0 / other
    def __rtruediv__(self: T, other: Any) -> None: raise DivisionByZeroError()


Empty = VBAEmpty()
