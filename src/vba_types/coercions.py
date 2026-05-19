from __future__ import annotations
from .exception import TypeMismatchError, InvalidNullError
from .types_registry import registry
from .null import VBANull


def type_error(value: VBATypeBase) -> None:
    raise TypeMismatchError()


def invalid_null(value: VBATypeBase) -> None:
    raise InvalidNullError()


def return_self(value: VBATypeBase) -> VBATypeBase:
    return value

registry.register(("*", VBANull), invalid_null)
registry.register(("Array", VBANull), type_error)
