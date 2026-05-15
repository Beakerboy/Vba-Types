from .vba_type_base import VBATypeBase


class IntegralType(VBATypeBase):
    
    def __int__(self: T) -> int:
        return self.value
