from typing import Any, Optional

class VBANothing:
    """
    Represents the VBA 'Nothing' type. 
    Used to indicate an uninitialized object variable.
    """
    _instance: Optional['VBANothing'] = None

    def __new__(cls) -> 'VBANothing':
        if cls._instance is None:
            cls._instance = super(VBANothing, cls).__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "Nothing"

    def __str__(self) -> str:
        # In VBA, printing Nothing usually results in an error, 
        # but for Python interoperability, we return its repr or empty.
        return "Nothing"

    def __bool__(self) -> bool:
        return False

    def __eq__(self, other: Any) -> bool:
        # 'Nothing' only equals itself. 
        # In VBA, you'd usually use 'If obj Is Nothing'
        return isinstance(other, VBANothing)

    # Disable math/string operations to mimic VBA's "Object not set" errors
    def __add__(self, other): raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __sub__(self, other): raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __mul__(self, other): raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __truediv__(self, other): raise TypeError("VBA Error 91: Object variable or With block variable not set")
