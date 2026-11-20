from typing import TypeVar
from .integral_type import VBAIntegralType


T = TypeVar("T", bound="VBABoolean")


class VBABoolean(VBAIntegralType):
    MIN_VALUE: int = -1
    MAX_VALUE: int = 0
    value: int

    def __init__(self: T, value: bool = False) -> None:
        self.value = -1 if value else 0

    def __bool__(self: T) -> bool:
        return self.value == -1

    def vba_and(self: T, other: T) -> T:
        return type(self)(bool(self) and bool(other))
