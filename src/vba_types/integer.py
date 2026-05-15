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

    def __add__(self: T, other: VBATypeBase) -> VBATypeBase:
        if (
                isinstance(other, VBAInteger) or
                isinstance(other, vba_types.VBABoolean)
        ):
            return type(self)(self.value + int(other))
        else:
            return other + self

    def __sub__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value - other.value)

    def __mod__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value % int(other))

    def __mul__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value * int(other))

    def __rmul__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value * int(other))

    def __pow__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value ** int(other))

    def __truediv__(self: T, other: VBATypeBase) -> vba_types.double.VBADouble:
        # VBA '/' always returns a Double (float in Python)
        if other.value == 0:
            raise DivisionByZeroError()
        return vba_types.double.VBADouble(self.value / other.value)

    def __floordiv__(self: T, other: VBATypeBase) -> T:
        # VBA '\' is integer division
        return type(self)(self.value // other.value)
