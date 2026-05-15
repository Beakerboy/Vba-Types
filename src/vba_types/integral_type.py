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

    def __int__(self: T) -> int:
        return self.value
