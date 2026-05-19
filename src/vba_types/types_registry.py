from __future__ import annotations
from typing import TypeVar
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBATypesRegistry')


class VBATypesRegistry:

    def __init__(self: T) -> None:
        self._registry: dict[tuple[str, type], Callable] = {}

    def register(self: T,
                 op: str,
                 incoming_cls: type,
                 handler: Callable) -> None:
        self._registry[(op, incoming_cls)] = handler

registry = VBATypesRegistry()
