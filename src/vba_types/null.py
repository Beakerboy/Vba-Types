from __future__ import annotations
from vba_types.vba_type_base import VBATypeBase
from typing import Optional, Type, TypeVar


T = TypeVar('T', bound='VBANull')


class VBANull(VBATypeBase):
    """
    Represents the VBA 'Null' type.
    Features 'Null Propagation': most operations with Null return Null.
    """
    _instance: Optional['VBANull'] = None

    def __new__(cls: Type[T]) -> 'VBANull':
        if cls._instance is None:
            cls._instance = super(VBANull, cls).__new__(cls)
        return cls._instance

    def __repr__(self: T) -> str:
        return "Null"

    def __str__(self: T) -> str:
        return "Null"

    def __bool__(self: T) -> bool:
        # In VBA, 'If Null Then' results in an error or False-like behavior
        # depending on context, but it is effectively falsy in Python.
        return False


Null = VBANull()
