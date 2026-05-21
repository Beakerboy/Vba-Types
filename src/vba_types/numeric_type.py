from __future__ import annotations
from .vba_type_base import VBATypeBase
from typing import TypeVar


T = TypeVar("T", bound="VBANumericType")


class VBANumericType(VBATypeBase):
    value: int | float

    def __int__(self: T) -> int:
        return int(self.value)
