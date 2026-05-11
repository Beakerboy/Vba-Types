import pytest
from vba_types.nothing import VBANothing, Nothing


def test_nothing_singleton() -> None:
    """Ensure Nothing is a singleton."""
    assert VBANothing() is Nothing
    assert id(VBANothing()) == id(Nothing)


def test_nothing_boolean() -> None:
    """Nothing should be falsy."""
    assert bool(Nothing) is False
    if Nothing:
        pytest.fail("Nothing should evaluate to False")


def test_nothing_equality() -> None:
    """Nothing should only equal itself."""
    assert Nothing == VBANothing()
    assert Nothing != 0
    assert Nothing != ""
    assert Nothing is not None


def test_nothing_errors() -> None:
    """
    Math operations on Nothing should raise TypeErrors (mimicking VBA Error 91).
    """
    with pytest.raises(TypeError, match="Object variable"):
        _ = Nothing + 1

    with pytest.raises(TypeError, match="Object variable"):
        _ = Nothing * 5


def test_nothing_repr() -> None:
    assert repr(Nothing) == "Nothing"
