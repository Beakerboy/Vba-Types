from __future__ import annotations
from typing import Any, Iterator, Type, TypeAlias, TypeVar
from .exceptions import SubscriptOutOfRangeError
from .vba_type_base import VBATypeBase


VBAArraySequence: TypeAlias = list[VBATypeBase] | list["VBAArraySequence"]
T = TypeVar('T', bound='VBAArray')


class VBAArray(VBATypeBase):
    def __init__(self: T, *args: VBATypeBase, base: int = 0) -> None:
        self._data: VBAArraySequence = list(args)
        ubound = base + len(args) - 1
        self._bounds: list[int | tuple[int, int]] = [(base, ubound)]

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

    @classmethod
    def initialize(cls: Type[T],
                   *args: int | tuple[int, int],
                   empty: VBATypeBase) -> T:
        data = list(args)
        if len(data) == 1 and isinstance(data[0], int):
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

    def lbound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][0]

    def ubound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][1]

    def __repr__(self: T) -> str:
        return f"<VBAArray: Bounds {self._bounds}>"
