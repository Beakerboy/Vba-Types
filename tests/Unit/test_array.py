import pytest
from vba_types.array import VBAArray
from vba_types.integral_type import VBAInteger
from vba_types.empty import Empty
from vba_types.exceptions import SubscriptOutOfRangeError


def test_base_0_initialization() -> None:
    """Tests standard comma-separated initialization with default Base 0."""
    arr = VBAArray("apple", "banana", "cherry")
    assert arr[0] == "apple"
    assert arr[1] == "banana"
    assert arr[2] == "cherry"
    assert VBAArray.lbound(arr).value == 0
    assert VBAArray.ubound(arr).value == 2


def test_base_1_initialization() -> None:
    """Tests comma-separated initialization with explicit Base 1."""
    arr = VBAArray(100, 200, 300, base=1)
    assert arr[1] == 100
    assert arr[2] == 200
    assert arr[3] == 300
    assert VBAArray.lbound(arr).value == 1
    assert VBAArray.ubound(arr).value == 3


def test_initialize_with_number() -> None:
    arr = VBAArray.initialize(3, empty=Empty)
    assert VBAArray.lbound(arr).value == 0
    assert VBAArray.ubound(arr).value == 3
    assert arr[0] is Empty
    assert arr[3] is Empty


def test_initialize_with_tuple() -> None:
    arr = VBAArray.initialize((0, 3), empty=Empty)
    assert VBAArray.lbound(arr).value == 0
    assert VBAArray.ubound(arr).value == 3
    assert arr[0] is Empty
    assert arr[3] is Empty


def test_multidimensional_custom_bounds() -> None:
    """Tests Array(1 To 2, 1 To 6) style initialization."""
    # Rows: 1 to 2, Cols: 1 to 6
    arr = VBAArray.initialize((1, 2), (1, 6), empty=Empty)

    # Set and Get
    arr[1, 1] = "Top-Left"
    arr[2, 6] = "Bottom-Right"
    arr[1, 4] = "Middleish"

    assert arr[1, 1] == "Top-Left"
    assert arr[2, 6] == "Bottom-Right"
    assert arr[1, 4] == "Middleish"

    # Check bounds
    assert VBAArray.lbound(arr, VBAInteger(1)).value == 1
    assert VBAArray.ubound(arr, VBAInteger(1)).value == 2
    assert VBAArray.lbound(arr, VBAInteger(2)).value == 1
    assert VBAArray.ubound(arr, VBAInteger(2)).value == 6


def test_out_of_bounds_raises_error() -> None:
    """
    Ensures that accessing indices outside the defined bounds raises
    an Exception.
    """
    arr = VBAArray(1, 2, 3, base=1)

    with pytest.raises(SubscriptOutOfRangeError):
        _ = arr[0]  # Too low for base 1

    with pytest.raises(SubscriptOutOfRangeError):
        _ = arr[4]  # Too high


def test_assignment_updates_value() -> None:
    """Verifies that __setitem__ actually modifies the internal data."""
    arr = VBAArray(None, None)
    arr[0] = "Modified"
    assert arr[0] == "Modified"
