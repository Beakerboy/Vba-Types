from __future__ import annotations
from .vba_type_base import VBATypeBase
from typing import TypeVar


T = TypeVar("T", bound="VBAIntegralType")


class VBAIntegralType(VBATypeBase):

    value: int

    def __init__(self: T, value: int = 0) -> None:
        self.value = self._validate(value)

    def _validate(self: T, value: int) -> int:
        if not (self.MIN_VALUE <= final_val <= self.MAX_VALUE):
            raise OverflowError("Run-time error '6': Overflow")
        return final_val
    
    def __int__(self: T) -> int:
        return self.value

    def __repr__(self: T) -> str:
        return str(self.value)
