import vba_types
from functools import total_ordering
from typing import TypeVar
from .exceptions import TypeMismatchError
from .null import Null
from .vba_type_base import VBATypeBase


T = TypeVar("T", bound="VBABoolean")


@total_ordering
class VBABoolean(VBATypeBase):
    value: int
    def __init__(self: T, value: bool = False) -> None:
        self.value = -1 if value else 0

    def __bool__(self: T) -> bool:
        return self.value == -1

    def __eq__(self: T, other: VBATypeBase) -> T | vba_types.null.VBANull:
        if other is Null:
             return Null
        if isinstance(other, vba_types.string.VBAString):
            if other.value.lower() == "true":
                return VBABoolean(self.value == -1)
            if other.value.lower() == "false":
                return VBABoolean(self.value == 0)
            raise TypeMismatchError()
        return VBABoolean(self.value == other.value)

    def __lt__(self: T, other: VBATypeBase) -> T | vba_types.null.VBANull:
        if other is Null:
             return Null
        if isinstance(other, vba_types.string.VBAString):
            if other.value.lower() == "true":
                return VBABoolean(self.value < -1)
            if other.value.lower() == "false":
                return VBABoolean(self.value < 0)
            raise TypeMismatchError()
        return VBABoolean(self.value < other.value)

    def vba_and(self: T, other: VBABoolean) -> VBABoolean:
        return VBABoolean(bool(self) and bool(other))
