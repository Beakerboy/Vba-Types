from __future__ import annotations
import pytest
from vba_types.variable import VBAVariable
from vba_types.integral_type import VBAInteger
from vba_types.string import VBAString
from vba_types.empty import Empty


def test_constructor() -> None:
    var = VBAVariable()
    assert var.declared_type == "variant"
    assert var.value is Empty


@pytest.mark.parametrize(
    "one, two, expected", [
        (
            VBAVariable("Integer", VBAInteger(10)),
            VBAVariable(value=VBAString()),
            True
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAVariable(value=VBAString()),
            True
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAVariable("String", VBAString()),
            False
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAVariable("String", VBAString("foo")),
            True
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAVariable("String", VBAString("0")),
            False
        ),
        (
            VBAVariable("Integer", VBAInteger(10)),
            VBAVariable(value=VBAString("0")),
            False
        ),
        (
            VBAVariable("Integer", VBAInteger(10)),
            VBAVariable("String", VBAString("0")),
            False
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAVariable(value=VBAString("0")),
            True
        ),
    ])
def test_variant_string_relation(one: VBAVariable,
                                 two: VBAVariable | VBAString,
                                 expected: bool) -> None:
    assert (one < two) == expected
