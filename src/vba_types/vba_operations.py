from collections.abc import Callable
from itertools import product
from .vba_registry import registry
from .vba_type_base import VBATypeBase
from .array import VBAArray
from .boolean import VBABoolean
from .double import VBADouble
from .empty import Empty, VBAEmpty
from .float_type import VBAFloatType
from .integral_type import (
    VBAIntegralType, VBAByte, VBAInteger, VBALong, VBALongLong
)
from .null import Null, VBANull
from .numeric_type import VBANumericType
from .string import VBAString
from .exceptions import DivisionByZeroError, TypeMismatchError


def register_multiple(operator_handlers: dict,
                      type_set1: list, type_set2: list) -> None:
    """
    Register all combinations of operators and types, and associate if the type
    are dissimilar
    """
    for op, handler in operator_handlers.items():
        for types in product(type_set1, type_set2):
            registry.register(op, types[0], types[1], handler)

            # Add reversed signature if different
            if len(types) == 2 and types[0] != types[1]:
                registry.register(op, types[1], types[0], handler)


def register_against_all(operators: tuple, type_: type[VBATypeBase],
                         handler: Callable) -> None:
    """
    Register a type with all types
    """
    for op in operators:
        registry.register(op, type_, VBATypeBase, handler)
        registry.register(op, VBATypeBase, type_, handler)


def handle_null_propogation(left: VBATypeBase, right: VBATypeBase) -> VBANull:
    return Null


def type_mismatch(left: VBATypeBase, right: VBATypeBase) -> None:
    raise TypeMismatchError


def _add_promote_to_byte(left: VBATypeBase,
                         right: VBATypeBase) -> VBAByte:
    return VBAByte(left.value + right.value)


def _sub_promote_to_byte(left: VBATypeBase,
                         right: VBATypeBase) -> VBAByte:
    return VBAByte(left.value - right.value)


def _mul_promote_to_byte(left: VBATypeBase,
                         right: VBATypeBase) -> VBAByte:
    return VBAByte(left.value * right.value)


def _add_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value + right.value)


def _sub_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value - right.value)


def _mul_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value * right.value)


def _add_promote_to_long(left: VBATypeBase,
                         right: VBATypeBase) -> VBALong:
    return VBALong(left.value + right.value)


def _add_promote_to_longlong(left: VBATypeBase,
                             right: VBATypeBase) -> VBALongLong:
    return VBALongLong(left.value + right.value)


def _sub_promote_to_longlong(left: VBATypeBase,
                             right: VBATypeBase) -> VBALongLong:
    return VBALongLong(left.value - right.value)


def _mul_promote_to_longlong(left: VBATypeBase,
                             right: VBATypeBase) -> VBALongLong:
    return VBALongLong(left.value * right.value)


def _mul_promote_to_long(left: VBATypeBase,
                         right: VBATypeBase) -> VBALong:
    return VBALong(left.value * right.value)


def _mod_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value % right.value)


def _mod_promote_to_long(left: VBATypeBase,
                         right: VBATypeBase) -> VBALong:
    return VBALong(left.value % right.value)


def _add_promote_to_double(left: VBATypeBase, right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value + right.value)


def _sub_promote_to_double(left: VBATypeBase, right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value - right.value)


def _sub_promote_to_long(left: VBATypeBase, right: VBATypeBase) -> VBALong:
    return VBALong(left.value - right.value)


def _mul_promote_to_double(left: VBATypeBase,
                           right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value * right.value)


def _pow_promote_to_double(left: VBATypeBase,
                           right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value ** right.value)


def _truediv_promote_to_double(left: VBATypeBase,
                               right: VBATypeBase) -> VBADouble:
    if right.value == 0.0 or right is Empty:
        raise DivisionByZeroError()
    return VBADouble(left.value / right.value)


def _floordiv_promote_to_long(left: VBATypeBase,
                              right: VBATypeBase) -> VBALong:
    if right.value == 0.0 or right is Empty:
        raise DivisionByZeroError()
    return VBALong(left.value // right.value)


def _floordiv_promote_to_integer(left: VBATypeBase,
                                 right: VBATypeBase) -> VBAInteger:
    if right.value == 0.0 or right is Empty:
        raise DivisionByZeroError()
    return VBAInteger(left.value // right.value)


def _numeric_equality(left: VBATypeBase,
                      right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value == right.value)


def _numeric_lt(left: VBATypeBase,
                right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value < right.value)


def _numeric_string_equality(left: VBATypeBase,
                             right: VBATypeBase) -> VBABoolean:
    return VBABoolean(float(left.value) == float(right.value))


def _numeric_string_lt(left: VBATypeBase,
                       right: VBATypeBase) -> VBABoolean:
    return VBABoolean(float(left.value) < float(right.value))


def _string_equality(left: VBATypeBase,
                     right: VBATypeBase) -> VBABoolean:
    return VBABoolean(str(left) == str(right))


def _string_lt(left: VBATypeBase,
               right: VBATypeBase) -> VBABoolean:
    return VBABoolean(str(left) < str(right))


def _bool_string_equality(left: VBATypeBase,
                          right: VBATypeBase) -> VBABoolean:
    if isinstance(left, VBAString):
        s = left
        b = right
    elif isinstance(right, VBAString):
        s = right
        b = left
    if s.value.lower() == "true":
        return VBABoolean(b.value == -1)
    if s.value.lower() == "false":
        return VBABoolean(b.value == 0)
    raise TypeMismatchError()


register_against_all(("+", "-", "*", "/", "//", "%", "==", "<", ">", "<>", "<=", "=>"),
                     VBANull, handle_null_propogation)
register_against_all(("+", "-", "*", "/", "//", "%"), VBAArray, type_mismatch)
register_multiple(
    {
        '+': _add_promote_to_byte,
        '-': _sub_promote_to_byte,
        '*': _mul_promote_to_byte,
    },
    [VBAByte, VBAEmpty],
    [VBAByte]
)
register_multiple(
    {
        '+': _add_promote_to_integer,
        '-': _sub_promote_to_integer,
        '*': _mul_promote_to_integer,
    },
    [VBABoolean, VBAInteger, VBAEmpty],
    [VBABoolean, VBAInteger]
)
register_multiple(
    {
        '+': _add_promote_to_integer,
        '-': _sub_promote_to_integer,
        '*': _mul_promote_to_integer,
        '/': _truediv_promote_to_double,
        '//': _floordiv_promote_to_integer,
        '%': _mod_promote_to_integer,
    }, [VBAEmpty], [VBAEmpty]
)
register_multiple(
    {
        '+': _add_promote_to_long,
        '-': _sub_promote_to_long,
        '*': _mul_promote_to_long,
    },
    [VBABoolean, VBAInteger, VBALong, VBAEmpty],
    [VBALong]
)
register_multiple(
    {
        '+': _add_promote_to_longlong,
        '-': _sub_promote_to_longlong,
        '*': _mul_promote_to_longlong,
    },
    [VBAIntegralType, VBAEmpty],
    [VBALongLong]
)
register_multiple(
    {
        '+': _add_promote_to_double,
        '-': _sub_promote_to_double,
        '*': _mul_promote_to_double,
        '/': _truediv_promote_to_double,
    },
    [VBAFloatType, VBAIntegralType, VBAEmpty],
    [VBADouble]
)

registry.register("%", VBAInteger, VBADouble, _mod_promote_to_integer)
registry.register("%", VBAInteger, VBAInteger, _mod_promote_to_long)
registry.register("*", VBAEmpty, VBADouble, _mul_promote_to_double)
registry.register("*", VBADouble, VBAEmpty, _mul_promote_to_double)

registry.register("/", VBANumericType,
                  VBANumericType, _truediv_promote_to_double)
registry.register("/", VBAIntegralType, VBAEmpty, _truediv_promote_to_double)
registry.register("/", VBAEmpty, VBAIntegralType, _truediv_promote_to_double)

registry.register("//", VBAIntegralType,
                  VBAIntegralType, _floordiv_promote_to_long)
registry.register("//", VBAIntegralType, VBAEmpty, _floordiv_promote_to_long)
registry.register("//", VBAEmpty, VBAIntegralType, _floordiv_promote_to_long)
registry.register("//", VBAIntegralType, VBADouble, _floordiv_promote_to_long)
registry.register("//", VBADouble, VBADouble, _floordiv_promote_to_long)

registry.register("**", VBANumericType, VBANumericType, _pow_promote_to_double)

registry.register("==", VBAArray, VBATypeBase, type_mismatch)
registry.register("==", VBATypeBase, VBAArray, type_mismatch)
registry.register("<", VBAArray, VBATypeBase, type_mismatch)
registry.register("<", VBATypeBase, VBAArray, type_mismatch)

registry.register("==", VBANumericType, VBANumericType, _numeric_equality)
registry.register("<", VBANumericType, VBANumericType, _numeric_lt)

registry.register("==", VBANumericType, VBAString, _numeric_string_equality)
registry.register("<", VBANumericType, VBAString, _numeric_string_lt)
registry.register("==", VBAString, VBANumericType, _numeric_string_equality)
registry.register("<", VBAString, VBANumericType, _numeric_string_lt)

registry.register("==", VBAEmpty, VBANumericType, _numeric_equality)
registry.register("<", VBAEmpty, VBANumericType, _numeric_lt)
registry.register("==", VBANumericType, VBAEmpty, _numeric_equality)
registry.register("<", VBANumericType, VBAEmpty, _numeric_lt)
registry.register("==", VBAEmpty, VBAEmpty, _numeric_equality)
registry.register("<", VBAEmpty, VBAEmpty, _numeric_lt)

registry.register("==", VBAEmpty, VBAString, _string_equality)
registry.register("==", VBAString, VBAEmpty, _string_equality)
registry.register("==", VBAString, VBAString, _string_equality)
registry.register("<", VBAString, VBAString, _string_lt)

registry.register("==", VBABoolean, VBAString, _bool_string_equality)
registry.register("==", VBAString, VBABoolean, _bool_string_equality)
