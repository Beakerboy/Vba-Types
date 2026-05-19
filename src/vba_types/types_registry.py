from __future__ import annotations
from typing import TypeVar
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBATypesRegistry')


class VBATypesRegistry:

    def coerce(self: T, incoming: VBATypeBase, declared_type: str) -> VBATypeBase:
        if declared_type == "Variant":
            return incoming


registry = VBATypesRegistry()
