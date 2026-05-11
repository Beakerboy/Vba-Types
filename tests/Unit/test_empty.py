import pytest
from vba_types.empty import VBAEmpty, Empty
from vba_types.exceptions import DivisionByZeroError


def test_singleton_behavior() -> None:
    """Ensure that only one instance of VBAEmpty ever exists."""
    another_empty = VBAEmpty()
    assert another_empty is Empty
    assert id(another_empty) == id(Empty)


def test_string_and_repr() -> None:
    """Test string representations."""
    assert str(Empty) == ""
    assert repr(Empty) == "Empty"


def test_type_casting() -> None:
    """Test explicit casting to Python primitives."""
    assert int(Empty) == 0
    assert float(Empty) == 0.0
    assert bool(Empty) is False


def test_equality_comparisons() -> None:
    """Test equality logic against various types."""
    assert Empty == VBAEmpty()  # Identity/Type
    assert Empty == 0           # Integer context
    assert Empty == 0.0         # Float context
    assert Empty == ""          # String context

    # Non-equal cases
    assert Empty != 1
    assert Empty != "0"
    assert Empty != [1, 2, 3]
    assert Empty is not None    # Empty is distinct from None


def test_arithmetic_operations() -> None:
    """Test that Empty behaves like 0 in math operations."""
    # Addition
    assert Empty + 10 == 10
    assert 5 + Empty == 5

    # Subtraction
    assert Empty - 5 == -5
    assert 10 - Empty == 10

    # Multiplication
    assert Empty * 100 == 0
    assert 50 * Empty == 0

    # Division
    assert Empty / 2 == 0.0
    with pytest.raises(DivisionByZeroError):
        10 / Empty


def test_boolean_context() -> None:
    """Test how Empty behaves in if-statements."""
    if Empty:
        pytest.fail("Empty should evaluate to False")
    assert not Empty
