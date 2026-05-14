from __future__ import annotations
from typing import Callable, Dict, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from vba_base import VBAValue

# Maps: (operator, left_type, right_type) -> calculation function
BinaryOpMap = Dict[Tuple[str, str, str], Callable[[Any, Any], 'VBAValue']]

class VBARegistry:
    def __init__(self):
        self._binary_ops: BinaryOpMap = {}
