from __future__ import annotations
from .exceptions import OverflowException
from .numeric_type import VBANumericType
from typing import TypeVar


T = TypeVar("T", bound="VBAIntegralType")


class VBAIntegralType(VBANumericType):

    MIN_VALUE: int
    MAX_VALUE: int
    value: int

    def __init__(self: T, value: int = 0) -> None:
        self.value = self._validate(value)

    def _validate(self: T, value: int) -> int:
        if not (self.MIN_VALUE <= value <= self.MAX_VALUE):
            raise OverflowException()
        return value

    def __int__(self: T) -> int:
        return self.value

    def __repr__(self: T) -> str:
        return str(self.value)

    def __index__(self: T) -> int:
        """Allows the object to be used in slice indices or bin() functions."""
        return self.value


class VBAInteger(VBAIntegralType):
    """
    Simulates the VBA Integer data type (16-bit signed).
    Range: -32,768 to 32,767.
    """
    MIN_VALUE: int = -32768
    MAX_VALUE: int = 32767
