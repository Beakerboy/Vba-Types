from vba_types.variable import VBAVariable
from vba_types.integer import VBAInteger
from vba_types.string import VBAString
from vba_types.empty import Empty


def test_constructor() -> None:
    var = VBAVariable()
    assert var.declared_type == "Variant"
    assert var.value is Empty


def test_variant_string_relation() -> None:
    foo = VBAVariable(value=VBAInteger(10))
    assert foo < VBAString()
