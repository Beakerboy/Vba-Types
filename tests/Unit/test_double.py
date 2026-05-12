import unittest
from .exceptions import DivisionByZeroError
from .vba_double import VBADouble


def test_initialization_and_value(self):
    """Test basic initialization with ints, floats, and implicit casting."""
    self.assertEqual(VBADouble(10).value, 10.0)
    self.assertEqual(VBADouble(10.5).value, 10.5)

    # Test initialization from another VBADouble instance
    vba_inner = VBADouble(5.5)
    self.assertEqual(VBADouble(vba_inner).value, 5.5)

  
def test_overflow_boundaries(self):
    """Test that values exceeding MAX_VALUE correctly raise OverflowError."""
    max_val = 1.7976931348623157e+308

    # Should work fine at the exact boundary
    self.assertEqual(VBADouble(max_val).value, max_val)
    self.assertEqual(VBADouble(-max_val).value, -max_val)

    # Should raise OverflowError when exceeding it
    with self.assertRaises(OverflowError):
        VBADouble(2e308)
    with self.assertRaises(OverflowError):
        VBADouble(-2e308)


def test_underflow_to_zero(self):
    """Test that tiny subnormal values snap down to 0.0 like VBA."""
    # Value smaller than 4.94065645841247e-324
    tiny_val = 1e-325
    self.assertEqual(VBADouble(tiny_val).value, 0.0)
    self.assertEqual(VBADouble(-tiny_val).value, 0.0)


def test_comparisons_and_type_safety(self):
    """Test comparisons, math.isclose equity, and safely handling invalid types."""
    v = VBADouble(5.5)

    # Standard comparisons
    self.assertTrue(v == 5.5)
    self.assertTrue(v > 5.0)
    self.assertTrue(v < 6.0)

    # Floating point precision safe check
    self.assertTrue(VBADouble(0.1 + 0.2) == 0.3)

    # Type safety: strings shouldn't crash the program with ValueError
    self.assertFalse(v == "not a float")

    # Less-than with unsupported types should return NotImplemented (TypeError in python)
    with self.assertRaises(TypeError):
        _ = v < "not a float"


def test_basic_math_operations(self):
    """Test regular math operations (+, -, *, /, **)."""
    v1 = VBADouble(10.5)

    self.assertEqual((v1 + 2.5).value, 13.0)
    self.assertEqual((v1 - 0.5).value, 10.0)
    self.assertEqual((v1 * 2).value, 21.0)
    self.assertEqual((v1 / 2).value, 5.25)
    self.assertEqual((VBADouble(2.0) ** 3).value, 8.0)


def test_reflected_math_operations(self):
    """Test right-side operations where native type is on the left."""
    v = VBADouble(4.0)

    self.assertEqual((10.0 + v).value, 14.0)
    self.assertEqual((10.0 - v).value, 6.0)
    self.assertEqual((2.0 * v).value, 8.0)
    self.assertEqual((12.0 / v).value, 3.0)
    self.assertEqual((2.0 ** v).value, 16.0)


def test_vba_integer_division(self):
    """Test that floor division mimics VBA's '\' behavior by dropping decimals first."""
    # VBA drops decimals before dividing: 9.9 becomes 9, 2.9 becomes 2 -> 9 // 2 = 4
    self.assertEqual((VBADouble(9.9) // VBADouble(2.9)).value, 4.0)

    # Check reflected floor division
    self.assertEqual((9.9 // VBADouble(2.9)).value, 4.0)


def test_division_by_zero(self):
    """Test that both standard and integer division throw Custom DivisionByZeroError."""
    v = VBADouble(5.0)
    zero = VBADouble(0.0)

    # True division checks
    with self.assertRaises(DivisionByZeroError):
        _ = v / 0
    with self.assertRaises(DivisionByZeroError):
        _ = v / zero
    with self.assertRaises(DivisionByZeroError):
        _ = 5.0 / zero

    # Floor division checks
    with self.assertRaises(DivisionByZeroError):
        _ = v // 0
    with self.assertRaises(DivisionByZeroError):
        _ = v // zero
    with self.assertRaises(DivisionByZeroError):
        _ = 5.0 // zero
