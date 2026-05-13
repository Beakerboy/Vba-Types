from functools import total_ordering
from typing import TypeVar
from .vba_type_base import VBATypeBase


T = TypeVar("T", bound="VBABoolean")


@total_ordering
class VBABoolean(VBATypeBase):
    value: int
    def __init__(self: T, value: bool = False) -> None:
        self.value = -1 if value else 0

    def __eq__(self: T, other: VBATypeBase) -> bool:
        if other is Null:
             return Null
        return self.value == other.value

    def __lt__(self: T, other: VBATypeBase) -> bool:
        if other is Null:
             return Null
        return self.value < other.value
