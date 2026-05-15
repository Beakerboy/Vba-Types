from typing import TypeVar
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBAString')


class VBAString(VBATypeBase):
    def __init__(self: T, value: str = "") -> None:
        self.value = value

    def __str__(self: T) -> str:
        return self.value

    def __add__(self: T, other: VBATypeBase) -> T:
        return VBAString(self.value + str(other.value))
