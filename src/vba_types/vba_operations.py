from .vba_registry import registry
from vba_types.vba_type_base import VBATypeBase
from vba_types.boolean import VBABoolean
from vba_types.double import VBADouble
from vba_types.empty import Empty, VBAEmpty
from vba_types.integral_type import VBAIntegralType, VBAInteger, VBALong
from vba_types.null import Null, VBANull
from vba_types.numeric_type import VBANumericType
from vba_types.string import VBAString
from .exceptions import DivisionByZeroError, TypeMismatchError


def handle_null_propogation(left: VBATypeBase, right: VBATypeBase) -> VBANull:
    return Null


def _add_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value + right.value)


def _sub_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value - right.value)


def _mul_promote_to_integer(left: VBATypeBase,
                            right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value * right.value)


def _add_promote_to_double(left: VBATypeBase, right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value + right.value)


def _sub_promote_to_double(left: VBATypeBase, right: VBATypeBase) -> VBADouble:
    return VBADouble(left.value - right.value)


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


def _numeric_equality(left: VBATypeBase,
                      right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value == right.value)


def _numeric_inequality(left: VBATypeBase,
                        right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value != right.value)


def _numeric_lt(left: VBATypeBase,
                right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value < right.value)


def _numeric_gt(left: VBATypeBase,
                right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value > right.value)


def _numeric_ge(left: VBATypeBase,
                right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value >= right.value)


def _numeric_le(left: VBATypeBase,
                right: VBATypeBase) -> VBABoolean:
    return VBABoolean(left.value <= right.value)


def _string_equality(left: VBATypeBase,
                     right: VBATypeBase) -> VBABoolean:
    return VBABoolean(str(left) == str(right))


def _string_inequality(left: VBATypeBase,
                       right: VBATypeBase) -> VBABoolean:
    return VBABoolean(str(left) != str(right))


def _bool_string_equality(left: VBATypeBase,
                          right: VBATypeBase) -> VBABoolean:
    if isinstance(left, VBAString):
        s = left
        b = right
    else:
        s = right
        b = left
    if s.value.lower() == "true":
        return VBABoolean(b.value == -1)
    if s.value.lower() == "false":
        return VBABoolean(b.value == 0)
    raise TypeMismatchError()


def _bool_string_inequality(left: VBATypeBase,
                            right: VBATypeBase) -> VBABoolean:
    if isinstance(left, VBAString):
        s = left
        b = right
    else:
        s = right
        b = left
    if s.value.lower() == "true":
        return VBABoolean(b.value != -1)
    if s.value.lower() == "false":
        return VBABoolean(b.value != 0)
    raise TypeMismatchError()


registry.register("+", VBANull, VBATypeBase, handle_null_propogation)
registry.register("+", VBATypeBase, VBANull, handle_null_propogation)
registry.register("-", VBANull, VBATypeBase, handle_null_propogation)
registry.register("-", VBATypeBase, VBANull, handle_null_propogation)
registry.register("*", VBANull, VBATypeBase, handle_null_propogation)
registry.register("*", VBATypeBase, VBANull, handle_null_propogation)
registry.register("/", VBANull, VBATypeBase, handle_null_propogation)
registry.register("/", VBATypeBase, VBANull, handle_null_propogation)
registry.register("//", VBANull, VBATypeBase, handle_null_propogation)
registry.register("//", VBATypeBase, VBANull, handle_null_propogation)
registry.register("%", VBANull, VBATypeBase, handle_null_propogation)
registry.register("%", VBATypeBase, VBANull, handle_null_propogation)

registry.register("+", VBAInteger, VBAInteger, _add_promote_to_integer)
registry.register("+", VBABoolean, VBAInteger, _add_promote_to_integer)
registry.register("+", VBAInteger, VBABoolean, _add_promote_to_integer)
registry.register("+", VBAInteger, VBAEmpty, _add_promote_to_integer)
registry.register("+", VBAEmpty, VBAInteger, _add_promote_to_integer)
registry.register("+", VBAEmpty, VBAEmpty, _add_promote_to_integer)
registry.register("-", VBAInteger, VBAInteger, _sub_promote_to_integer)
registry.register("-", VBABoolean, VBAInteger, _sub_promote_to_integer)
registry.register("-", VBAInteger, VBABoolean, _sub_promote_to_integer)
registry.register("-", VBAEmpty, VBAEmpty, _sub_promote_to_integer)
registry.register("-", VBAEmpty, VBAInteger, _sub_promote_to_integer)
registry.register("-", VBAInteger, VBAEmpty, _sub_promote_to_integer)

registry.register("+", VBADouble, VBADouble, _add_promote_to_double)
registry.register("+", VBAIntegralType, VBADouble, _add_promote_to_double)
registry.register("+", VBAEmpty, VBADouble, _add_promote_to_double)
registry.register("+", VBADouble, VBAIntegralType, _add_promote_to_double)
registry.register("+", VBADouble, VBAEmpty, _add_promote_to_double)
registry.register("-", VBADouble, VBADouble, _sub_promote_to_double)
registry.register("-", VBAIntegralType, VBADouble, _sub_promote_to_double)
registry.register("-", VBAEmpty, VBADouble, _sub_promote_to_double)
registry.register("-", VBADouble, VBAIntegralType, _sub_promote_to_double)
registry.register("-", VBADouble, VBAEmpty, _sub_promote_to_double)

registry.register("*", VBAInteger, VBAInteger, _mul_promote_to_integer)
registry.register("*", VBABoolean, VBAInteger, _mul_promote_to_integer)
registry.register("*", VBAInteger, VBABoolean, _mul_promote_to_integer)
registry.register("*", VBAInteger, VBAEmpty, _mul_promote_to_integer)
registry.register("*", VBAEmpty, VBAInteger, _mul_promote_to_integer)
registry.register("*", VBAEmpty, VBAEmpty, _mul_promote_to_integer)

registry.register("*", VBADouble, VBADouble, _mul_promote_to_double)
registry.register("*", VBAIntegralType, VBADouble, _mul_promote_to_double)
registry.register("*", VBAEmpty, VBADouble, _mul_promote_to_double)
registry.register("*", VBADouble, VBAIntegralType, _mul_promote_to_double)
registry.register("*", VBADouble, VBAEmpty, _mul_promote_to_double)

registry.register("/", VBAIntegralType,
                  VBAIntegralType, _truediv_promote_to_double)
registry.register("/", VBAIntegralType, VBAEmpty, _truediv_promote_to_double)
registry.register("/", VBAEmpty, VBAIntegralType, _truediv_promote_to_double)
registry.register("/", VBAIntegralType, VBADouble, _truediv_promote_to_double)
registry.register("/", VBADouble, VBADouble, _truediv_promote_to_double)

registry.register("//", VBAIntegralType,
                  VBAIntegralType, _floordiv_promote_to_long)
registry.register("//", VBAIntegralType, VBAEmpty, _floordiv_promote_to_long)
registry.register("//", VBAEmpty, VBAIntegralType, _floordiv_promote_to_long)
registry.register("//", VBAIntegralType, VBADouble, _floordiv_promote_to_long)
registry.register("//", VBADouble, VBADouble, _floordiv_promote_to_long)

registry.register("**", VBANumericType, VBANumericType, _pow_promote_to_double)

registry.register("==", VBANull, VBATypeBase, handle_null_propogation)
registry.register("==", VBATypeBase, VBANull, handle_null_propogation)
registry.register("<>", VBANull, VBATypeBase, handle_null_propogation)
registry.register("<>", VBATypeBase, VBANull, handle_null_propogation)
registry.register("=>", VBANull, VBATypeBase, handle_null_propogation)
registry.register("=>", VBATypeBase, VBANull, handle_null_propogation)
registry.register("<=", VBANull, VBATypeBase, handle_null_propogation)
registry.register("<=", VBATypeBase, VBANull, handle_null_propogation)
registry.register("<", VBANull, VBATypeBase, handle_null_propogation)
registry.register("<", VBATypeBase, VBANull, handle_null_propogation)
registry.register(">", VBANull, VBATypeBase, handle_null_propogation)
registry.register(">", VBATypeBase, VBANull, handle_null_propogation)

registry.register("==", VBANumericType, VBANumericType, _numeric_equality)
registry.register("<>", VBANumericType, VBANumericType, _numeric_inequality)
registry.register(">", VBANumericType, VBANumericType, _numeric_gt)
registry.register("<", VBANumericType, VBANumericType, _numeric_lt)
registry.register("<=", VBANumericType, VBANumericType, _numeric_le)
registry.register("=>", VBANumericType, VBANumericType, _numeric_ge)

registry.register("==", VBAEmpty, VBANumericType, _numeric_equality)
registry.register("<>", VBAEmpty, VBANumericType, _numeric_inequality)
registry.register(">", VBAEmpty, VBANumericType, _numeric_gt)
registry.register("<", VBAEmpty, VBANumericType, _numeric_lt)
registry.register("<=", VBAEmpty, VBANumericType, _numeric_le)
registry.register("=>", VBAEmpty, VBANumericType, _numeric_ge)
registry.register("==", VBANumericType, VBAEmpty, _numeric_equality)
registry.register("<>", VBANumericType, VBAEmpty, _numeric_inequality)
registry.register(">", VBANumericType, VBAEmpty, _numeric_gt)
registry.register("<", VBANumericType, VBAEmpty, _numeric_lt)
registry.register("<=", VBANumericType, VBAEmpty, _numeric_le)
registry.register("=>", VBANumericType, VBAEmpty, _numeric_ge)
registry.register("==", VBAEmpty, VBAEmpty, _numeric_equality)
registry.register("<>", VBAEmpty, VBAEmpty, _numeric_inequality)
registry.register(">", VBAEmpty, VBAEmpty, _numeric_gt)
registry.register("<", VBAEmpty, VBAEmpty, _numeric_lt)
registry.register("<=", VBAEmpty, VBAEmpty, _numeric_le)
registry.register("=>", VBAEmpty, VBAEmpty, _numeric_ge)

registry.register("==", VBAEmpty, VBAString, _string_equality)
registry.register("==", VBAString, VBAEmpty, _string_equality)
registry.register("==", VBAString, VBAString, _string_equality)
registry.register("<>", VBAString, VBAString, _string_inequality)
registry.register("<>", VBAEmpty, VBAString, _string_inequality)
registry.register("<>", VBAString, VBAEmpty, _string_inequality)

registry.register("==", VBABoolean, VBAString, _bool_string_equality)
registry.register("==", VBAString, VBABoolean, _bool_string_equality)
registry.register("<>", VBABoolean, VBAString, _bool_string_inequality)
registry.register("<>", VBAString, VBABoolean, _bool_string_inequality)
