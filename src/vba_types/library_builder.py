from __future__ import annotations
import inspect
import types
from typing import Any, Callable, Type, TypeVar, Union, get_args, get_origin
from .contracts import (
    LibModuleDefinition,
    LibProjectDefinition,
    LibraryDefinition,
    ParamDefinition,
    PassMech,
)
from .boolean import VBABoolean
from .double import VBADouble
from .empty import Empty
from .integral_type import VBAInteger, VBALong
from .numeric_type import VBANumericType
from .string import VBAString
from .vba_type_base import VBATypeBase
from .variable import VBAVariable

_TYPE_MAP: dict[type[VBATypeBase], str] = {
    VBABoolean: "boolean",
    VBAInteger: "integer",
    VBALong: "long",
    VBADouble: "double",
    VBAString: "string",
    VBANumericType: "variant",
    VBATypeBase: "variant",
}


def vba_proc(
    *,
    byref: list[str] | None = None,
    name: str | None = None,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """
    Optional decorator to override VBA procedure name or ByRef parameters.
    """
    def decorator(func: Callable[..., Any]) -> Callable[..., Any]:
        setattr(func, "_vba_byref", set(byref or []))
        if name is not None:
            setattr(func, "_vba_name", name.lower())
        return func
    return decorator


T = TypeVar("T", bound="LibraryBuilder")


class LibraryBuilder:
    def __init__(self: T, name: str) -> None:
        self.name = name.lower()
        self._modules: dict[str, LibModuleDefinition] = {}
        self._classes: dict[str, LibModuleDefinition] = {}

    def add_module(self: T, module_name: str, cls: Type[Any]) -> LibraryBuilder:
        mod_key = module_name.lower()
        self._modules[mod_key] = {
            "name": mod_key,
            "type": "Module",
            "functions": self._extract_procedures(mod_key, cls),
        }
        return self

    def add_class(self, class_name: str, cls: Type[Any]) -> LibraryBuilder:
        cls_key = class_name.lower()
        self._classes[cls_key] = {
            "name": cls_key,
            "type": "Class",
            "functions": self._extract_procedures(cls_key, cls),
        }
        return self

    def build(self) -> LibProjectDefinition:
        return {
            "name": self.name,
            "type": "Library",
            "modules": self._modules,
            "classes": self._classes,
        }

    def _extract_procedures(
        self: T, module_name: str, cls: Type[Any]
    ) -> dict[str, LibraryDefinition]:
        procedures: dict[str, LibraryDefinition] = {}

        for attr_name, raw_member in cls.__dict__.items():
            if attr_name.startswith("_"):
                continue

            # Unwrap @staticmethod / @classmethod to inspect the underlying
            # function
            if isinstance(raw_member, (staticmethod, classmethod)):
                func = raw_member.__func__
            elif callable(raw_member):
                func = raw_member
            else:
                continue

            handle = getattr(cls, attr_name)
            proc_name = getattr(func, "_vba_name", attr_name.lower())
            byref_params: set[str] = getattr(func, "_vba_byref", set())

            sig = inspect.signature(func, eval_str=True)
            is_sub = (
                sig.return_annotation is None or 
                sig.return_annotation is inspect.Signature.empty
            )

            params: list[ParamDefinition] = []
            for param_name, param in sig.parameters.items():
                if param_name in ("self", "cls"):
                    continue

                optional = param.default is not inspect.Parameter.empty
                default_val = param.default if optional else Empty
                mech: PassMech = "ByVal"
                if param_name in byref_params 
                    mech  = "ByRef"

                params.append({
                    "name": param_name.lower(),
                    "optional": optional,
                    "default": default_val,
                    "var": self._make_vba_variable(param.annotation),
                    "mech": mech,
                })

            procedures[proc_name] = {
                "name": proc_name,
                "type": "Sub" if is_sub else "Function",
                "project": self.name,
                "module": module_name,
                "handle": handle,
                "params": params,
            }

        return procedures

    @staticmethod
    def _unwrap_optional(annotation: Any) -> Any:
        """Unwraps Optional[T] / T | None to T."""
        origin = get_origin(annotation)
        if origin is Union or origin is types.UnionType:
            args = [a for a in get_args(annotation) if a is not type(None)]
            if len(args) == 1:
                return args[0]
        return annotation

    @classmethod
    def _make_vba_variable(cls: Type[T], annotation: Any) -> VBAVariable:
        resolved = cls._unwrap_optional(annotation)
        if isinstance(resolved, type) and issubclass(resolved, VBATypeBase):
            declared = _TYPE_MAP.get(resolved, "variant")
            return VBAVariable(declared_type=declared)
        return VBAVariable(declared_type="variant")
