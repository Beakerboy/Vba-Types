from __future__ import annotations
from typing import TypeVar
from .vba_type_base import VBATypeBase


T = TypeVar('T', bound='VBATypesRegistry')


class VBATypesRegistry:

    def __init__(self: T) -> None:
        self._registry: dict[tuple[str, type], Callable] = {}

    def register(self: T,
                 type_name: str,
                 incoming_cls: type,
                 handler: Callable) -> None:
        self._registry[(type_name, incoming_cls)] = handler

    def coerce(self: T,
                type_name: str,
                incoming: VBATypeBase) -> VBATypeBase:
        handler = self._get_handler(type_name, type(incoming))

        if not handler:
            raise TypeError(self._get_vba_error_msg(op, left, right))
        return handler(incoming)

registry = VBATypesRegistry()
