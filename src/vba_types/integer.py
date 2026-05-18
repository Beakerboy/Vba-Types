from __future__ import annotations
from .exceptions import DivisionByZeroError
from .boolean import VBABoolean
from .null import Null
from .integral_type import VBAIntegralType
import vba_types
from typing import TypeVar, TYPE_CHECKING


if TYPE_CHECKING:
    from vba_types.vba_type_base import VBATypeBase


T = TypeVar("T", bound="VBAInteger")


class VBAInteger(VBAIntegralType):
    """
    Simulates the VBA Integer data type (16-bit signed).
    Range: -32,768 to 32,767.
    """
    MIN_VALUE: int = -32768
    MAX_VALUE: int = 32767

    def __index__(self: T) -> int:
        """Allows the object to be used in slice indices or bin() functions."""
        return self.value

    def __eq__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value == other.value)

    def __ne__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value != other.value)

    def __lt__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value < other.value)

    def __le__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value <= other.value)

    def __ge__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value >= other.value)

    def __gt__(self: T, other: VBATypeBase) -> VBABoolean:
        if other is Null:
            return Null
        return VBABoolean(self.value > other.value)
