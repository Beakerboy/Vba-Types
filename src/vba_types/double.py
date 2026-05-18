from __future__ import annotations
import math
from typing import TypeVar, TYPE_CHECKING
from .float_type import VBAFloatType


T = TypeVar("T", bound="VBADouble")


class VBADouble(VBAFloatType):
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

    def __init__(self: T, value: 0.0) -> None:
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

        # Handle Underflow: VBA rounds numbers closer to 0 than MIN_POSITIVE
        # down to 0.0
        if 0.0 < abs(value) < self.MIN_POSITIVE:
            return 0.0

        return value

    def __repr__(self: T) -> str:
        return str(self.value)

    def __float__(self: T) -> float:
        return self.value
