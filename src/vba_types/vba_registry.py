from __future__ import annotations
from typing import (
    Any, Callable, Dict, Tuple, TypeVar, TYPE_CHECKING
)


if TYPE_CHECKING:
    from vba_base import VBAValue


T = TypeVar('T', bound='VBARegistry')


# Maps: (operator, left_type, right_type) -> calculation function
BinaryOpMap = Dict[Tuple[str, str, str], Callable[[Any, Any], 'VBAValue']]


class VBARegistry:
    def __init__(self: T):
        self._binary_ops: BinaryOpMap = {}

    def register_binary(self: T, op: str, type_left: str, type_right: str, handler: Callable):
        self._binary_ops[(op, type_left, type_right)] = handler

    def execute_binary_op(self: T, op: str, left: VBAValue, right: VBAValue) -> VBAValue:
        key = (op, left.type_name, right.type_name)
        handler = self._binary_ops.get(key)
        
        if not handler:
            raise TypeError(self._get_vba_error_msg(op, left, right))
        return handler(left, right)

    def _get_vba_error_msg(self: T,
                           op: str,
                           left: VBAValue,
                           right: VBAValue) -> str:
        return (f"Run-time error '13': Type mismatch for "
                f"{left.type_name} {op} {right.type_name}")


registry = VBARegistry()
