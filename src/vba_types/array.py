from __future__ import annotations
from typing import Any, Iterator, Tuple, Type, TypeVar, Union
from .exceptions import SubscriptOutOfRangeError
from .exceptions import TypeMismatchError
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBAArray')


@total_ordering
class VBAArray(VBATypeBase):
    def __init__(self: T, *args: VBATypeBase, base: int = 0) -> None:
        self._data = list(args)
        self._bounds = [(base, base + len(args) - 1)]

    def __getitem__(self: T, key: Union[int, Tuple[int, ...]]) -> Any:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        val = self._data
        for c in coords:
            val = val[c]
        return val

    def __setitem__(self: T,
                    key: Union[int, Tuple[int, ...]],
                    value: Any) -> None:
        indices = key if isinstance(key, tuple) else (key,)
        coords = self._get_coords(indices)
        target = self._data
        for c in coords[:-1]:
            target = target[c]
        target[coords[-1]] = value

    def __iter__(self: T) -> Iterator[Any]:
        for item in self._data:
            yield item

    @classmethod
    def initialize(cls: Type[T],
                   *args: int | list[tuple[int, int]],
                   empty: VBATypeBase) -> T:
        data = list(args)
        if len(data) == 1 and not isinstance(data[0], tuple):
            input = [empty] * (data[0] + 1)
            return cls(*input)
        else:
            arr = cls.__new__(cls)
            arr._bounds = list(args)
            shape = tuple(
                max_idx - min_idx + 1 for min_idx, max_idx in arr._bounds
            )
            arr._data = arr._recursive_init(shape, empty)
            return arr

    def _recursive_init(self: T,
                        shape: Tuple[int, ...],
                        empty: VBATypeBase) -> Any:
        if len(shape) == 1:
            return [empty] * shape[0]
        rng = range(shape[0])
        return [self._recursive_init(shape[1:], empty) for _ in rng]

    def _get_coords(self: T, indices: Tuple[int, ...]) -> Tuple[int, ...]:
        if len(indices) != len(self._bounds):
            raise SubscriptOutOfRangeError()

        internal = []
        for i, idx in enumerate(indices):
            low, high = self._bounds[i]
            if not (low <= idx <= high):
                raise SubscriptOutOfRangeError()
            internal.append(idx - low)
        return tuple(internal)

    def lbound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][0]

    def ubound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][1]

    def __repr__(self: T) -> str:
        return f"<VBAArray: Bounds {self._bounds}>"
