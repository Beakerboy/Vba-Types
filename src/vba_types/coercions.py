from __future__ import annotations
from .exceptions import TypeMismatchError, InvalidNullError
from .types_registry import registry
from .empty import VBAEmpty
from .integral_type import VBAInteger
from .null import VBANull
from .vba_type_base import VBATypeBase


def type_error(value: VBATypeBase) -> None:
    raise TypeMismatchError()


def invalid_null(value: VBATypeBase) -> None:
    raise InvalidNullError()


def return_self(value: VBATypeBase) -> VBATypeBase:
    return value


def zero(value: VBATypeBase) -> VBAInteger:
    return VBAInteger()


registry.register("*", VBANull, invalid_null)
registry.register("array", VBANull, type_error)
registry.register("integer", VBAEmpty, zero)
