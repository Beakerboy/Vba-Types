import math
from functools import total_ordering
from typing import Union, TypeVar
from .exceptions import DivisionByZeroError
from .vba_type_base import VBATypeBase


# Type alias for types that can interact with VBADouble
VBACompatible = Union[int, float, "VBATypeBase"]
T = TypeVar("T", bound="VBADouble")


@total_ordering
class VBADouble(VBATypeBase):
    """
    Simulates the VBA Double data type (64-bit floating-point).
    Negative range: -1.7976931348623157E+308 to -4.94065645841247E-324
    Positive range: 4.94065645841247E-324 to 1.7976931348623157E+308
    """
    # Max value for IEEE 754 double precision
    MAX_VALUE: float = 1.7976931348623157e+308
    # Smallest positive subnormal value
    MIN_POSITIVE: float = 4.94065645841247e-324

    value: float

    def __init__(self: T, value: VBACompatible = 0.0) -> None:
        # Avoid double validation if we are already dealing with a verified
        # float
        if isinstance(value, float):
            raw_val = value
        elif hasattr(value, "value"):
            raw_val = float(value.value)
        else:
            try:
                raw_val = float(value)  # type: ignore
            except (ValueError, TypeError):
                return NotImplemented

        self.value = self._validate(raw_val)

    def _validate(self: T, value: float) -> float:
        # Handle Overflow: check if value exceeds absolute maximum limits
        # Python floats turn into 'inf' if they exceed the 64-bit limit during
        # math
        if math.isinf(value) or abs(value) > self.MAX_VALUE:
            raise OverflowError("Run-time error '6': Overflow")

        # Handle Underflow: VBA rounds numbers closer to 0 than MIN_POSITIVE down
        # to 0.0
        if 0.0 < abs(value) < self.MIN_POSITIVE:
            return 0.0

        return value

    def __repr__(self: T) -> str:
        return str(self.value)

    def __float__(self: T) -> float:
        return self.value

    def __int__(self: T) -> int:
        # VBA converts Double to Integer/Long using Banker's rounding.
        # However, for a generic __int__, casting via a helper math pattern is
        # safer.
        return int(self.value)

    def _safefloat(self: T, other: VBACompatible) -> float:
        """Helper to extract a float safely or raise a TypeError."""
        if hasattr(other, "value"):
            return float(other.value)
        return float(other)  # type: ignore

    # --- Comparison Operators ---
    def __eq__(self: T, other: object) -> bool:
        try:
            return math.isclose(self.value, self._safefloat(other))
        except (TypeError, ValueError):
            return False

    def __lt__(self: T, other: VBACompatible) -> bool:
        try:
            return self.value < self._safefloat(other)
        except (TypeError, ValueError):
            return NotImplemented

    # --- Math Operators ---
    def __add__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value + self._safefloat(other))

    def __radd__(self: T, other: VBACompatible) -> T:
        return type(self)(self._safefloat(other) + self.value)

    def __sub__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value - self._safefloat(other))

    def __rsub__(self: T, other: VBACompatible) -> T:
        return type(self)(self._safefloat(other) - self.value)

    def __mul__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value * self._safefloat(other))

    def __rmul__(self: T, other: VBACompatible) -> T:
        return type(self)(self._safefloat(other) * self.value)

    def __pow__(self: T, other: VBACompatible) -> T:
        return type(self)(self.value ** self._safefloat(other))

    def __rpow__(self: T, other: VBACompatible) -> T:
        return type(self)(self._safefloat(other) ** self.value)

    def __truediv__(self: T, other: VBACompatible) -> T:
        # VBA '/' operator on Doubles returns a Double
        denom = self._safefloat(other)
        if denom == 0.0:
            raise DivisionByZeroError()
        return type(self)(self.value / denom)

    def __rtruediv__(self: T, other: VBACompatible) -> T:
        if self.value == 0.0:
            raise DivisionByZeroError()
        return type(self)(self._safefloat(other) / self.value)

    def __floordiv__(self: T, other: VBACompatible) -> T:
        # VBA '\' (Integer division) drops decimal points BEFORE dividing. This
        # converts operands to integers first via Banker's Rounding, then
        # divides. To replicate VBA's '\' fully, you would typically return a
        # VBA Long/Integer. For this class, we cast both to integers first.
        denom = int(self._safefloat(other))
        if denom == 0:
            raise DivisionByZeroError()
        return type(self)(int(self.value) // denom)

    def __rfloordiv__(self: T, other: VBACompatible) -> T:
        if int(self.value) == 0:
            raise DivisionByZeroError()
        return type(self)(int(self._safefloat(other)) // int(self.value))
