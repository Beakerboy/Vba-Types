import pytest
from vba_types.variable import VBAVariable
from vba_types.integer import VBAInteger
from vba_types.string import VBAString
from vba_types.empty import Empty


def test_constructor() -> None:
    var = VBAVariable()
    assert var.declared_type == "Variant"
    assert var.value is Empty


@pytest.mark.parametrize(
    "one, two, expected", [
        (
            VBAVariable("Integer", VBAInteger(10))),
            VBAVariable("String", VBAString()),
            True
        ),
        (
            VBAVariable(value=VBAInteger(10)),
            VBAString(),
            True
        ),    
])
def test_variant_string_relation(one, two, expected: bool) -> None:
    assert (one < two) == expected
