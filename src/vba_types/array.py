from typing import Any, Tuple, Type, TypeVar, Union


T = TypeVar('T', bound='VBAArray')


class VBAArray:
    def __init__(self: T, *args: Any, base: int = 0) -> None:
        self._data = list(args)
        self._bounds = [(base, base + len(args) - 1)]

    @classmethod
    def initialize(cls, *args: int | list[tuple[int, int]], empty: Any = None)-> Type[T]:
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
            arr._data = arr._recursive_init(shape)
            return arr

    def _recursive_init(self: T, shape: Tuple[int, ...]) -> Any:
        if len(shape) == 1:
            return [None] * shape[0]
        return [self._recursive_init(shape[1:]) for _ in range(shape[0])]

    def _get_coords(self: T, indices: Tuple[int, ...]) -> Tuple[int, ...]:
        if len(indices) != len(self._bounds):
            raise IndexError("Subscript out of range (dimension mismatch)")
        
        internal = []
        for i, idx in enumerate(indices):
            low, high = self._bounds[i]
            if not (low <= idx <= high):
                raise IndexError(f"Subscript out of range: {idx} (Expected {low} to {high})")
            internal.append(idx - low)
        return tuple(internal)

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

    def lbound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][0]

    def ubound(self: T, dimension: int = 1) -> int:
        return self._bounds[dimension - 1][1]

    def __repr__(self: T) -> str:
        return f"<VBAArray: Bounds {self._bounds}>"
