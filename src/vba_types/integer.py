import vba_types
from functools import total_ordering
from typing import Union, TypeVar
from .exceptions import DivisionByZeroError
from .null import Null
from .vba_type_base import VBATypeBase


T = TypeVar("T", bound="VBAInteger")


@total_ordering
class VBAInteger(VBATypeBase):
    """
    Simulates the VBA Integer data type (16-bit signed).
    Range: -32,768 to 32,767.
    """
    MIN_VALUE: int = -32768
    MAX_VALUE: int = 32767
    value: int

    def __init__(self: T, value: VBATypeBase = 0) -> None:
        self.value = self._validate(value)

    def _validate(self: T, value: VBATypeBase) -> int:
        # Extract raw numeric value
        if isinstance(value, VBAInteger):
            raw_val = float(value.value)
        else:
            raw_val = float(value)

        # VBA uses 'Banker's Rounding'
        # (rounds to nearest even number)
        final_val: int = int(round(raw_val))

        if not (self.MIN_VALUE <= final_val <= self.MAX_VALUE):
            raise OverflowError("Run-time error '6': Overflow")
        return final_val

    def __repr__(self: T) -> str:
        return str(self.value)

    def __int__(self: T) -> int:
        return self.value

    def __index__(self: T) -> int:
        """Allows the object to be used in slice indices or bin() functions."""
        return self.value

    def __eq__(self: T, other: VBATypeBase) -> bool:
        if other is Null:
             return Null
        return self.value == other.value

    def __lt__(self: T, other: VBATypeBase) -> bool:
        if other is Null:
             return Null
        return self.value < int(other)

    def __add__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value + int(other))

    def __radd__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value + int(other))

    def __sub__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value - int(other))

    def __rsub__(self: T, other: VBATypeBase) -> T:
        return type(self)(int(other) - self.value)

    def __mod__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value % int(other))

    def __mul__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value * int(other))

    def __rmul__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value * int(other))

    def __pow__(self: T, other: VBATypeBase) -> T:
        return type(self)(self.value ** int(other))

    def __truediv__(self: T, other: VBATypeBase) -> float:
        # VBA '/' always returns a Double (float in Python)
        if other == 0:
            raise DivisionByZeroError()
        return vba_types.double.VBADouble(self.value / other.value)

    def __floordiv__(self: T, other: VBATypeBase) -> T:
        # VBA '\' is integer division
        return type(self)(self.value // other.value)
