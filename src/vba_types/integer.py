from typing import Union, TypeVar

# Type alias for types that can be converted to/interact with VBAInteger
VBACompatible = Union[int, float, "VBAInteger"]
T = TypeVar("T", bound="VBAInteger")

class VBAInteger:
    """
    Simulates the VBA Integer data type (16-bit signed).
    Range: -32,768 to 32,767.
    """
    MIN_VALUE: int = -32768
    MAX_VALUE: int = 32767
    value: int

    def __init__(self, value: VBACompatible = 0) -> None:
        self.value = self._validate(value)

    def _validate(self, value: VBACompatible) -> int:
        # Extract raw numeric value
        raw_val: float = float(value.value) if isinstance(value, VBAInteger) else float(value)
        
        # VBA uses 'Banker's Rounding' (rounds to nearest even number)
        final_val: int = int(round(raw_val))
        
        if not (self.MIN_VALUE <= final_val <= self.MAX_VALUE):
            raise OverflowError("Run-time error '6': Overflow")
        return final_val

    def __repr__(self) -> str:
        return str(self.value)

    def __int__(self) -> int:
        return self.value

    def __index__(self) -> int:
        """Allows the object to be used in slice indices or bin() functions."""
        return self.value

    def __add__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value + int(other))

    def __radd__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value + int(other))

    def __sub__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value - int(other))

    def __mul__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value * int(other))

    def __truediv__(self, other: VBACompatible) -> float:
        # VBA '/' always returns a Double (float in Python)
        return float(self.value) / float(other)

    def __floordiv__(self: T, other: VBACompatible) -> T:
        # VBA '\' is integer division
        return type(self)(self.value // int(other))
