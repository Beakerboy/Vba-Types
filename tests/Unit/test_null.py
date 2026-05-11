import pytest
from vba_types.null import VBANull, Null


def test_null_singleton() -> None:
    assert VBANull() is Null
    assert id(VBANull()) == id(Null)


def test_null_propagation_arithmetic() -> None:
    """Any math operation with Null should return Null."""
    assert Null + 5 is Null
    assert 10 - Null is Null
    assert Null * 2 is Null
    assert 100 / Null is Null
    assert (Null + 10) * 5 is Null


def test_vba_comparison_behavior() -> None:
    """In VBA, Null = Null and Null = 0 are both NOT True."""
    assert (Null == Null) is False
    assert (Null == 0) is False
    assert (Null == "") is False
    assert (Null != 10) is True


def test_null_boolean() -> None:
    """Null should be falsy."""
    assert bool(Null) is False
