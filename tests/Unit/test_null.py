from vba_types.null import VBANull, Null
from vba_types.integer import VBAInteger
from vba_types.string import VBAString


def test_null_singleton() -> None:
    assert VBANull() is Null
    assert id(VBANull()) == id(Null)


def test_null_propagation_arithmetic() -> None:
    """Any math operation with Null should return Null."""
    assert Null + VBAInteger(5) is Null
    assert VBAInteger(10) - Null is Null
    assert Null * VBAInteger(2) is Null
    assert VBAInteger(100) / Null is Null
    assert (Null + VBAInteger(10)) * VBAInteger(5) is Null


def test_vba_comparison_behavior() -> None:
    assert (Null == Null) is Null
    assert (Null == VBAInteger(0)) is Null
    assert (Null == VBAString("")) is Null
    assert (Null != VBAInteger(10)) is Null


def test_null_boolean() -> None:
    """Null should be falsy."""
    assert bool(Null) is False
