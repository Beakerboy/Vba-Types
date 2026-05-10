from typing import Any, Optional, Union

class VBAEmpty:
    """
    Represents the VBA 'Empty' type. 
    It is initialized to 0 in a numeric context and "" in a string context.
    """
    _instance: Optional['VBAEmpty'] = None

    def __new__(cls) -> 'VBAEmpty':
        if cls._instance is None:
            cls._instance = super(VBAEmpty, cls).__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "Empty"

    def __str__(self) -> str:
        return ""

    def __int__(self) -> int:
        return 0

    def __float__(self) -> float:
        return 0.0

    def __bool__(self) -> bool:
        # In VBA, Empty evaluates to False/0
        return False

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, VBAEmpty):
            return True
        if isinstance(other, (int, float)):
            return other == 0
        if isinstance(other, str):
            return other == ""
        return False

    # Arithmetic behavior (Empty acts as 0)
    def __add__(self, other: Any) -> Any: return 0 + other
    def __radd__(self, other: Any) -> Any: return other + 0
    def __sub__(self, other: Any) -> Any: return 0 - other
    def __rsub__(self, other: Any) -> Any: return other - 0
    def __mul__(self, other: Any) -> Any: return 0 * other
    def __rmul__(self, other: Any) -> Any: return other * 0
    def __truediv__(self, other: Any) -> Any: return 0 / other


Empty = VBAEmpty()
