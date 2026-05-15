from .vba_type_base import VBATypeBase


class IntegralType(VBATypeBase):
    def __init__(self: T, value: int = 0) -> None:
        self.value = self._validate(value)

    def _validate(self: T, value: VBATypeBase) -> int:
        # Extract raw numeric value
        if isinstance(value, VBAInteger):
            raw_val = value.value
        else:
            raw_val = int(value)
        # VBA uses 'Banker's Rounding'
        # (rounds to nearest even number)
        final_val: int = int(round(raw_val))

        if not (self.MIN_VALUE <= final_val <= self.MAX_VALUE):
            raise OverflowError("Run-time error '6': Overflow")
        return final_val
    
    def __int__(self: T) -> int:
        return self.value
