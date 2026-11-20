from __future__ import annotations
from typing import Any, Iterator, Type, TypeAlias, TypeVar
from .exceptions import SubscriptOutOfRangeError
from .vba_type_base import VBATypeBase


VBAArraySequence: TypeAlias = list[VBATypeBase] | list["VBAArraySequence"]
VBAArrayData: TypeAlias = VBATypeBase | list["VBAArrayData"]
T = TypeVar('T', bound='VBAArray')


class VBAArray(VBATypeBase):
    def __init__(self: T, *args: VBATypeBase, base: int = 0) -> None:
        self._data: VBAArraySequence = list(args)
        ubound = base + len(args) - 1
        self._bounds = [(base, ubound)]

    def __getitem__(self: T, key: int | tuple[int, ...]) -> VBATypeBase:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        val: Any = self._data
        for c in coords:
            val = val[c]
        if isinstance(val, list):
            raise SubscriptOutOfRangeError("Not enough indices provided.")
        output: VBATypeBase = val
        return output

    def __setitem__(self: T,
                    key: int | tuple[int, ...],
                    value: Any) -> None:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        target: Any = self._data
        for c in coords[:-1]:
            target = target[c]
        if not isinstance(target, list):
            raise SubscriptOutOfRangeError("Too many indices provided.")
        target[coords[-1]] = value

    def __iter__(self: T) -> Iterator[Any]:
        for item in self._data:
            yield item

    def __repr__(self: T) -> str:
        return f"<VBAArray: Bounds {self._bounds}>"

    @classmethod
    def initialize(cls: Type[T],
                   *args: int | tuple[int, int],
                   empty: VBATypeBase) -> T:
        data = list(args)
        if isinstance(data[0], int):
            input = [empty] * (data[0] + 1)
            return cls(*input)
        else:
            arr = cls.__new__(cls)
            # A simple loop filters out any lingering type doubts for mypy
            arr._bounds = [item for item in args if isinstance(item, tuple)]
            shape = tuple(
                max_idx - min_idx + 1 for min_idx, max_idx in arr._bounds
            )
            arr._data = arr._recursive_init(shape, empty)
            return arr

    def lbound(self: T, dimension: VBATypeBase | None = None) -> VBATypeBase:
        from .integral_type import VBALong
        if dimension is None:
            dimension = VBALong(1)
        return VBALong(self._bounds[dimension.value - 1][0])

    def ubound(self: T, dimension: VBATypeBase | None = None) -> VBATypeBase:
        from .integral_type import VBALong
        if dimension is None:
            dimension = VBALong(1)
        return VBALong(self._bounds[dimension.value - 1][1])

    def redim(self: T, *bounds: tuple[int, int], preserve: bool = False,
              empty: VBATypeBase) -> None:
        """
        Implements VBA ReDim and ReDim Preserve.

        Assumes declaration legality has already been validated by
        the interpreter/compiler.
        """

        new_shape = tuple(
            upper - lower + 1
            for lower, upper in bounds
        )

        # ReDim Preserve validation
        if preserve:

            if len(bounds) != len(self._bounds):
                raise SubscriptOutOfRangeError()

            # All dimensions except the final one must be identical
            for old, new in zip(self._bounds[:-1], bounds[:-1]):
                if old != new:
                    raise SubscriptOutOfRangeError()

            # Last dimension: lower bound must remain the same
            old_low, old_high = self._bounds[-1]
            new_low, new_high = bounds[-1]

            if old_low != new_low:
                raise SubscriptOutOfRangeError()

            # Upper bound may change freely

        # Allocate new storage
        new_data = self._recursive_init(new_shape, empty)

        # Copy overlapping values
        if preserve:

            old_shape = tuple(
                upper - lower + 1
                for lower, upper in self._bounds
            )

            self._copy_overlap(
                self._data,
                new_data,
                old_shape,
                new_shape,
            )

        self._bounds = list(bounds)
        self._data = new_data

    # Check the behavior of altering an element in an initialized array.
    # [empty] * shape[0] may copy a reference to the empty object shape[0]
    # number of times, such that editing arr(0) would change it in all entries.
    # Do we need a deep clone or read the type of empty and initialize it.
    def _recursive_init(self: T,
                        shape: tuple[int, ...],
                        empty: VBATypeBase) -> Any:
        if len(shape) == 1:
            return [empty] * shape[0]
        rng = range(shape[0])
        return [self._recursive_init(shape[1:], empty) for _ in rng]

    def _get_coords(self: T, indices: tuple[int, ...]) -> tuple[int, ...]:
        if len(indices) != len(self._bounds):
            raise SubscriptOutOfRangeError()

        internal = []
        for i, idx in enumerate(indices):
            low, high = self._bounds[i]
            if not (low <= idx <= high):
                raise SubscriptOutOfRangeError()
            internal.append(idx - low)
        return tuple(internal)

    def _copy_overlap(self: T, src: VBAArrayData, dst: VBAArrayData,
                      old_shape: tuple[int, ...],
                      new_shape: tuple[int, ...]) -> None:
        """
        Recursively copies values from one array storage structure into
        another, preserving only the region that exists in both arrays.

        This method is used to implement VBA's `ReDim Preserve` behavior. The
        destination array may be either larger or smaller than the source
        array. Only elements whose indices are valid in both arrays are copied.
        Elements outside the overlapping region are either discarded (when
        shrinking) or retain their default initialization values (when
        expanding).

        The method operates on the internal zero-based list representation
        rather than VBA bounds. Bound validation should already have been
        performed by `redim()` before this method is called.

        Args:
            src: The source array data structure being resized from.
            dst: The newly allocated destination array data structure.
            old_shape: The dimensions of the source array, expressed as element
                counts for each dimension.
            new_shape: The dimensions of the destination array, expressed as
                element counts for each dimension.

        Examples:
            Expanding a one-dimensional array:

                [1, 2, 3] -> [1, 2, 3, 0, 0]

            Shrinking a one-dimensional array:

                [1, 2, 3, 4] -> [1, 2]

            Expanding the last dimension of a two-dimensional array:

                (3 x 2) -> (3 x 4)

            In each case, only the intersection of the source and destination
            arrays is copied.
        """
        limit = min(old_shape[0], new_shape[0])

        if len(old_shape) == 1:
            for i in range(limit):
                dst[i] = src[i]
            return

        assert isinstance(src, list)
        assert isinstance(dst, list)
        for i in range(limit):
            self._copy_overlap(
                src[i],
                dst[i],
                old_shape[1:],
                new_shape[1:],
            )
