from __future__ import annotations
from .boolean import VBABoolean
from .double import VBADouble
from .empty import VBAEmpty
from .exceptions import TypeMismatchError, InvalidNullError
from .integral_type import VBAInteger, VBALong
from .null import VBANull
from .numeric_type import VBANumericType
from .string import VBAString
from .types_registry import registry
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


def lon_self(value: VBANumericType) -> VBALong:
    return VBALong(int(value))


def dbl_self(value: VBANumericType) -> VBADouble:
    return VBADouble(float(value.value))


def bool_self(value: VBATypeBase) -> VBABoolean:
    return VBABoolean(bool(value))


registry.register("*", VBANull, invalid_null)
registry.register("array", VBANull, type_error)
registry.register("boolean", VBAEmpty, bool_self)
registry.register("boolean", VBABoolean, return_self)
registry.register("double", VBAEmpty, dbl_self)
registry.register("double", VBAInteger, dbl_self)
registry.register("double", VBALong, dbl_self)
registry.register("double", VBADouble, return_self)
registry.register("integer", VBAEmpty, int_self)
registry.register("integer", VBAInteger, return_self)
registry.register("integer", VBAString, int_self)
registry.register("long", VBAEmpty, lon_self)
registry.register("long", VBALong, return_self)
registry.register("string", VBAEmpty, str_self)
registry.register("string", VBAInteger, str_self)
registry.register("string", VBAString, return_self)
