from __future__ import annotations
from .exceptions import OverflowException
from .vba_type_base import VBATypeBase
from typing import TypeVar


T = TypeVar("T", bound="VBAIntegralType")


class VBAIntegralType(VBATypeBase):

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

    def __eq__(self: T, other: Any) -> VBABoolean:
        from .boolean import VBABoolean
        return VBABoolean(self.value == other.value)
