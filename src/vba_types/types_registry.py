from __future__ import annotations
from .vba_type_base import VBATypeBase


class VBATypesRegistry:

    def coerce(incoming: VBATypeBase, declared_type: str) -> VBATypeBase:
        pass


registry = VBATypesRegistry()
