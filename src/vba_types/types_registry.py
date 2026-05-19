from __future__ import annotations
from typing import Callable, TypeVar
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
            raise TypeError(self._get_vba_error_msg(type_name, incoming))
        return handler(incoming)

     def _get_handler(self: T,
                     declared_name: str,
                     incoming_cls: type) -> Optional[Callable]:
        exact_key = (declared_name, incoming_cls)
        if exact_key in self._registry:
            return self._registry[exact_key]
        exact_key = ("*", incoming_cls)
        if star_key in self._registry:
                return self._registry[star_key]
        return None

    def _get_vba_error_msg(self: T,
                           type_name: str,
                           incoming: VBATypeBase) -> str:
        return (f"Run-time error '13': Type mismatch for "
                f"{type_name} {type(incoming)}")


registry = VBATypesRegistry()
