from __future__ import annotations
from typing import Any, Callable, Literal, NotRequired, TypedDict
from .variable import VBAVariable

PassMech = Literal["ByVal", "ByRef"]


class ParamDefinition(TypedDict):
    name: str
    optional: bool
    default: Any
    var: VBAVariable
    mech: PassMech


class FunctionBase(TypedDict):
    type: str
    project: str
    module: str
    extra: NotRequired[dict[str, Any]]


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
