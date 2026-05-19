from vba_types.variable import VBAVariable
from vba_types.empty import Empty


def test_constructor() -> None:
    var = VBAVariable()
    assert var.declared_type == "Variant"
    assert var.value is Empty
