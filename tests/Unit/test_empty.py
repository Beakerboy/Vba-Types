import pytest
import vba_types
from vba_types.empty import VBAEmpty, Empty
from vba_types.integral_type import VBAInteger
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
    assert Empty == VBAInteger(0)           # Integer context
    assert Empty == vba_types.double.VBADouble(0.0)         # Float context
    assert Empty == vba_types.string.VBAString("")          # String context

    # Non-equal cases
    assert Empty != VBAInteger(1)
    assert Empty != vba_types.string.VBAString("0")
    assert Empty is not None    # Empty is distinct from None


def test_arithmetic_operations() -> None:
    """Test that Empty behaves like 0 in math operations."""
    # Addition
    int_ten = VBAInteger(10)
    assert Empty + int_ten == int_ten
    result = VBAInteger(5) + Empty
    assert result == VBAInteger(5)

    # Subtraction
    result = Empty - VBAInteger(5)
    assert result == VBAInteger(-5)
    assert int_ten - Empty == int_ten

    # Multiplication
    int_zero = VBAInteger(0)
    assert Empty * VBAInteger(100) == int_zero
    assert VBAInteger(50) * Empty == int_zero

    # Division
    result = Empty / VBAInteger(2)
    assert result == vba_types.double.VBADouble(0.0)
    with pytest.raises(DivisionByZeroError):
        int_ten / Empty


def test_boolean_context() -> None:
    """Test how Empty behaves in if-statements."""
    if Empty:
        pytest.fail("Empty should evaluate to False")
    assert not Empty
