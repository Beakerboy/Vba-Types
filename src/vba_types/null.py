from typing import Any, Optional

class VBANull:
    """
    Represents the VBA 'Null' type.
    Features 'Null Propagation': most operations with Null return Null.
    """
    _instance: Optional['VBANull'] = None

    def __new__(cls) -> 'VBANull':
        if cls._instance is None:
            cls._instance = super(VBANull, cls).__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "Null"

    def __str__(self) -> str:
        return "Null"

    def __bool__(self) -> bool:
        # In VBA, 'If Null Then' results in an error or False-like behavior 
        # depending on context, but it is effectively falsy in Python.
        return False

    # Null Propagation: Any math with Null returns Null
    def __add__(self, other: Any) -> 'VBANull': return self
    def __radd__(self, other: Any) -> 'VBANull': return self
    def __sub__(self, other: Any) -> 'VBANull': return self
    def __rsub__(self, other: Any) -> 'VBANull': return self
    def __mul__(self, other: Any) -> 'VBANull': return self
    def __rmul__(self, other: Any) -> 'VBANull': return self
    def __truediv__(self, other: Any) -> 'VBANull': return self
    def __rtruediv__(self, other: Any) -> 'VBANull': return self

    # VBA Comparison logic: Any comparison with Null returns Null
    # Note: In Python, __eq__ must return a boolean for dict/set keys to work.
    # To mimic VBA's "If x = Null" always being False, we return False for equality.
    def __eq__(self, other: Any) -> bool:
        return False  # In VBA, Null = Null is False (use IsNull instead)

    def __ne__(self, other: Any) -> bool:
        return True # In VBA, Null <> anything is also effectively handled as not-equal
