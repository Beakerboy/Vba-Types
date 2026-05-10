import pytest

def test_nothing_singleton():
    """Ensure Nothing is a singleton."""
    assert VBANothing() is Nothing
    assert id(VBANothing()) == id(Nothing)

def test_nothing_boolean():
    """Nothing should be falsy."""
    assert bool(Nothing) is False
    if Nothing:
        pytest.fail("Nothing should evaluate to False")

def test_nothing_equality():
    """Nothing should only equal itself."""
    assert Nothing == VBANothing()
    assert Nothing != 0
    assert Nothing != ""
    assert Nothing is not None

def test_nothing_errors():
    """Math operations on Nothing should raise TypeErrors (mimicking VBA Error 91)."""
    with pytest.raises(TypeError, match="Object variable"):
        _ = Nothing + 1
    
    with pytest.raises(TypeError, match="Object variable"):
        _ = Nothing * 5

def test_nothing_repr():
    assert repr(Nothing) == "Nothing"
