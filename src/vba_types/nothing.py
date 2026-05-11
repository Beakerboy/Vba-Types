from typing import Any, Optional, TypeVar


T = TypeVar('T', bound='VBANothing')


class VBANothing:
    """
    Represents the VBA 'Nothing' type.
    Used to indicate an uninitialized object variable.
    """
    _instance: Optional['VBANothing'] = None

    def __new__(cls: Type[T]) -> 'VBANothing':
        if cls._instance is None:
            cls._instance = super(VBANothing, cls).__new__(cls)
        return cls._instance

    def __repr__(self: T) -> str:
        return "Nothing"

    def __str__(self: T) -> str:
        # In VBA, printing Nothing usually results in an error,
        # but for Python interoperability, we return its repr or empty.
        return "Nothing"

    def __bool__(self: T) -> bool:
        return False

    def __eq__(self: T, other: Any) -> bool:
        # 'Nothing' only equals itself.
        # In VBA, you'd usually use 'If obj Is Nothing'
        return isinstance(other, VBANothing)

    # Disable math/string operations to mimic VBA's "Object not set" errors
    def __add__(self: T, other: Any) -> None:
        raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __sub__(self: T, other: Any) -> None:
        raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __mul__(self: T, other: Any) -> None:
        raise TypeError("VBA Error 91: Object variable or With block variable not set")
    def __truediv__(self: T, other: Any) -> None:
        raise TypeError("VBA Error 91: Object variable or With block variable not set")


Nothing = VBANothing()
