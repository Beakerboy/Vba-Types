import pytest
from vba_types.exceptions import DivisionByZeroError
from vba_types.double import VBADouble


def test_initialization_and_value() -> None:
    """Test basic initialization with ints, floats, and implicit casting."""
    assert VBADouble(10).value == 10.0
    assert VBADouble(10.5).value == 10.5

    # Test initialization from another VBADouble instance
    vba_inner = VBADouble(5.5)
    assert VBADouble(vba_inner).value == 5.5


def test_overflow_boundaries() -> None:
    """Test that values exceeding MAX_VALUE correctly raise OverflowError."""
    max_val = 1.7976931348623157e+308

    # Should work fine at the exact boundary
    assert VBADouble(max_val).value == max_val
    assert VBADouble(-max_val).value == -max_val

    # Should raise OverflowError when exceeding it
    with pytest.raises(OverflowError):
        VBADouble(2e308)
    with pytest.raises(OverflowError):
        VBADouble(-2e308)


def test_underflow_to_zero() -> None:
    """Test that tiny subnormal values snap down to 0.0 like VBA."""
    # Value smaller than 4.94065645841247e-324
    tiny_val = 1e-325
    assert VBADouble(tiny_val).value == 0.0
    assert VBADouble(-tiny_val).value == 0.0


def test_basic_math_operations() -> None:
    """Test regular math operations (+, -, *, /, **)."""
    v1 = VBADouble(10.5)
    v1 = VBADouble(1.5)

    assert (v1 + v2).value == 12.0
    assert (v1 - v2).value == 9.0
    assert (v1 * v2).value == 15.75
    assert (v1 / 2).value == 7.0
    assert (VBADouble(2.0) ** VBADouble(3)).value == 8.0


def test_vba_integer_division() -> None:
    """
    Test that floor division mimics VBA's '\' behavior by dropping decimals
    first.
    """
    # VBA drops decimals before dividing:
    # 9.9 becomes 9, 2.9 becomes 2 -> 9 // 2 = 4
    assert (VBADouble(9.9) // VBADouble(2.9)).value == 4.0

    # Check reflected floor division
    assert (9.9 // VBADouble(2.9)).value == 4.0


def test_division_by_zero() -> None:
    """
    Test that both standard and integer division throw Custom
    DivisionByZeroError.
    """
    v = VBADouble(5.0)
    zero = VBADouble(0.0)

    # True division checks
    with pytest.raises(DivisionByZeroError):
        _ = v / 0
    with pytest.raises(DivisionByZeroError):
        _ = v / zero
    with pytest.raises(DivisionByZeroError):
        _ = 5.0 / zero

    # Floor division checks
    with pytest.raises(DivisionByZeroError):
        _ = v // 0
    with pytest.raises(DivisionByZeroError):
        _ = v // zero
    with pytest.raises(DivisionByZeroError):
        _ = 5.0 // zero
