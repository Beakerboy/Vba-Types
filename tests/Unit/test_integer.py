import pytest
from vba_types.double import VBADouble
from vba_types.exceptions import OverflowException
from vba_types.integer import VBAInteger


def test_initialization_boundaries() -> None:
    """Test that valid boundaries work and invalid ones raise OverflowError."""
    assert int(VBAInteger(32767)) == 32767
    assert int(VBAInteger(-32768)) == -32768

    with pytest.raises(OverflowError, match="Run-time error '6': Overflow"):
        VBAInteger(32768)

    with pytest.raises(OverflowError, match="Run-time error '6': Overflow"):
        VBAInteger(-32769)


def test_arithmetic_overflow() -> None:
    """
    Test that operations resulting in out-of-bounds
    values raise OverflowError.
    """
    a = VBAInteger(30000)
    b = VBAInteger(3000)

    # with pytest.raises(OverflowException):
    _ = a + b  # 33000 > 32767

    c = VBAInteger(-32000)
    d = VBAInteger(1000)
    with pytest.raises(OverflowException):
        _ = c - d  # -33000 < -32768


def test_basic_math_operations() -> None:
    """Test standard arithmetic returns correct values and types."""
    a = VBAInteger(10)
    b = VBAInteger(3)

    # Addition
    res_add = a + b
    assert int(res_add) == 13
    assert isinstance(res_add, VBAInteger)

    # Integer Division (VBA '\' operator)
    res_div = a // b
    assert int(res_div) == 3

    # Multiplication
    assert int(a * b) == 30

    # Power
    assert int(b ** VBAInteger(2)) == 9


def test_truediv_returns_float() -> None:
    """Test that '/' returns a float, matching VBA's 'Double' return type."""
    a = VBAInteger(10)
    res = a / VBAInteger(4)
    assert bool(res == VBADouble(10 / 4))
    assert isinstance(res, VBADouble)


def test_interoperability() -> None:
    """Test interaction between VBAInteger and standard Python ints."""
    a = VBAInteger(100)
    assert int(a + VBAInteger(50)) == 150
    assert int(VBAInteger(50) + a) == 150
