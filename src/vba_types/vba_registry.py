from __future__ import annotations
from typing import (
    Any, Callable, Optional, TypeVar, TYPE_CHECKING
)


if TYPE_CHECKING:
    from vba_types.vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBARegistry')


# Maps: (operator, left_type, right_type) -> calculation function
BinaryOpMap = dict[tuple[str, str, str], Callable[[Any, Any], 'VBATypeBase']]


class VBARegistry:
    def __init__(self: T) -> None:
        self._registry: dict[tuple[str, type, type], Callable] = {}

    def register(self: T,
                 op: str,
                 left_cls: type,
                 right_cls: type,
                 handler: Callable) -> None:
        self._registry[(op, left_cls, right_cls)] = handler

    def _get_handler(self: T,
                     op: str,
                     left_cls: type,
                     right_cls: type) -> Optional[Callable]:
        exact_key = (op, left_cls, right_cls)
        if exact_key in self._registry:
            return self._registry[exact_key]
        for (reg_op, reg_left, reg_right), handler in self._registry.items():
            if (
                    reg_op == op and
                    issubclass(left_cls, reg_left) and
                    issubclass(right_cls, reg_right)
               ):
                return handler
        return None

    def execute(self: T,
                op: str,
                left: VBATypeBase,
                right: VBATypeBase) -> VBATypeBase:
        key = (op, type(left), type(right))
        handler = self._get_handler(key)

        if not handler:
            raise TypeError(self._get_vba_error_msg(op, left, right))
        return handler(left, right)

    def _get_vba_error_msg(self: T,
                           op: str,
                           left: VBATypeBase,
                           right: VBATypeBase) -> str:
        return (f"Run-time error '13': Type mismatch for "
                f"{type(left)} {op} {type(right)}")


registry = VBARegistry()
