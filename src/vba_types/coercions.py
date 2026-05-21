from __future__ import annotations
from .exceptions import TypeMismatchError, InvalidNullError
from .types_registry import registry
from .empty import VBAEmpty
from .integral_type import VBAInteger, VBALong
from .null import VBANull
from .numeric_type import VBANumericType
from .string import VBAString
from .vba_type_base import VBATypeBase


def type_error(value: VBATypeBase) -> None:
    raise TypeMismatchError()


def invalid_null(value: VBATypeBase) -> None:
    raise InvalidNullError()


def return_self(value: VBATypeBase) -> VBATypeBase:
    return value


def str_self(value: VBATypeBase) -> VBATypeBase:
    return VBAString(str(value))


def int_self(value: VBANumericType) -> VBAInteger:
    return VBAInteger(int(value))


def lon_self(value: VBATypeBase) -> VBALong:
    return VBALong(int(value))


registry.register("*", VBANull, invalid_null)
registry.register("array", VBANull, type_error)
registry.register("integer", VBAEmpty, int_self)
registry.register("integer", VBAString, int_self)
registry.register("long", VBAEmpty, lon_self)
registry.register("string", VBAEmpty, str_self)
registry.register("string", VBAInteger, str_self)
registry.register("integer", VBAInteger, return_self)
registry.register("long", VBALong, return_self)
registry.register("string", VBAString, return_self)
