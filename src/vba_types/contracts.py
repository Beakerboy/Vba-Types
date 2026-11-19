from __future__ import annotations
from typing import Any, Callable, TypedDict, TypeVar
from vba_types import VBAVariable


class ParamDefinition(TypedDict):
    name: str
    optional: bool
    default: Any
    var: VBAVariable
    mech: str


class FunctionBase(TypedDict):
    type: str
    project: str
    module: str
    extra: dict[str, Any]


class LibraryDefinition(FunctionBase):
    name: str
    handle: Callable[..., Any]
    params: list[ParamDefinition]


class LibModuleDefinition(TypedDict):
    name: str
    type: str
    functions: dict[str, LibraryDefinition]


class LibProjectDefinition(TypedDict):
    name: str
    type: str
    modules: dict[str, LibModuleDefinition]
    classes: dict[str, LibModuleDefinition]
