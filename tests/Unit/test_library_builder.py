from __future__ import annotations

from typing import Any, Optional, Type
import pytest

from vba_types.empty import Empty
from vba_types.integral_type import VBAInteger, VBALong
from vba_types.numeric_type import VBANumericType
from vba_types.library_builder import LibraryBuilder, vba_proc
from vba_types.string import VBAString
from vba_types.vba_type_base import VBATypeBase
from vba_types.variable import VBAVariable


class SampleModule:
    CONSTANT_ATTR: int = 42

    def _private_helper(self: SampleModule) -> None:
        pass

    @staticmethod
    def my_func(a: VBAInteger, b: VBALong = VBALong(10)) -> VBAString:
        return VBAString("ok")

    @classmethod
    def class_sub(cls: Type[SampleModule], flag: VBAInteger) -> None:
        pass

    def untyped_sub(self: SampleModule, x) -> None:  # noqa: ANN001
        pass

    @vba_proc(byref=["target"], name="CustomProcName")
    def original_name(
        self: SampleModule, target: VBALong, source: VBAString
    ) -> VBAInteger:
        return VBAInteger(1)

    @staticmethod
    def variadic_func(*args: VBANumericType) -> VBALong:
        return VBALong(0)


class SampleClass:
    def method_one(
        self: SampleClass,
        opt_pipe: VBAInteger | None = None,
        opt_typing: Optional[VBAString] = None,
        generic_val: VBATypeBase = Empty,
    ) -> VBALong:
        return VBALong(1)


class TestLibraryBuilder:
    def test_build_project_structure_and_lowercasing(
        self: TestLibraryBuilder,
    ) -> None:
        project = (
            LibraryBuilder("MyProject")
            .add_module("MyModule", SampleModule)
            .add_class("MyClass", SampleClass)
            .build()
        )

        assert project["name"] == "myproject"
        assert project["type"] == "library"
        assert "mymodule" in project["modules"]
        assert "myclass" in project["classes"]

        mod = project["modules"]["mymodule"]
        assert mod["name"] == "mymodule"
        assert mod["type"] == "module"

        cls_def = project["classes"]["myclass"]
        assert cls_def["name"] == "myclass"
        assert cls_def["type"] == "class"

    def test_ignores_private_and_non_callable_attributes(
        self: TestLibraryBuilder,
    ) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        funcs = project["modules"]["mod"]["functions"]

        assert "_private_helper" not in funcs
        assert "constant_attr" not in funcs
        assert "CONSTANT_ATTR" not in funcs

    def test_function_vs_sub_detection(self: TestLibraryBuilder) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        funcs = project["modules"]["mod"]["functions"]

        assert funcs["my_func"]["type"] == "Function"
        assert funcs["class_sub"]["type"] == "Sub"
        assert funcs["untyped_sub"]["type"] == "Sub"

    def test_staticmethod_and_classmethod_unwrapping(
        self: TestLibraryBuilder,
    ) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        funcs = project["modules"]["mod"]["functions"]

        assert [p["name"] for p in funcs["my_func"]["params"]] == ["a", "b"]
        assert [p["name"] for p in funcs["class_sub"]["params"]] == ["flag"]
        assert [p["name"] for p in funcs["untyped_sub"]["params"]] == ["x"]

        assert funcs["my_func"]["handle"] is SampleModule.my_func
        assert funcs["class_sub"]["handle"] == SampleModule.class_sub

    def test_required_and_optional_parameters(
        self: TestLibraryBuilder,
    ) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        params = project["modules"]["mod"]["functions"]["my_func"]["params"]

        assert params[0]["name"] == "a"
        assert params[0]["optional"] is False
        assert params[0]["default"] is Empty
        assert params[0]["mech"] == "ByVal"
        assert params[0]["var"].declared_type == "integer"

        assert params[1]["name"] == "b"
        assert params[1]["optional"] is True
        default_b = params[1]["default"]
        assert isinstance(default_b, VBALong)
        assert default_b.value == 10
        assert params[1]["mech"] == "ByVal"
        assert params[1]["var"].declared_type == "long"

    def test_vba_proc_decorator_overrides(self: TestLibraryBuilder) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        funcs = project["modules"]["mod"]["functions"]

        assert "original_name" not in funcs
        assert "customprocname" in funcs

        proc = funcs["customprocname"]
        assert proc["name"] == "customprocname"
        assert proc["type"] == "Function"

        params = proc["params"]
        assert params[0]["name"] == "target"
        assert params[0]["mech"] == "ByRef"
        assert params[0]["var"].declared_type == "long"

        assert params[1]["name"] == "source"
        assert params[1]["mech"] == "ByVal"
        assert params[1]["var"].declared_type == "string"

    def test_var_positional_args_expansion(self: TestLibraryBuilder) -> None:
        project = (
            LibraryBuilder("TestLib").add_module("Mod", SampleModule).build()
        )
        params = (
            project["modules"]["mod"]["functions"]["variadic_func"]["params"]
        )

        assert len(params) == 30
        assert params[0]["name"] == "arg1"
        assert params[0]["optional"] is False
        assert params[0]["default"] is Empty
        assert params[0]["var"].declared_type == "variant"

        for idx in range(1, 30):
            assert params[idx]["name"] == f"arg{idx + 1}"
            assert params[idx]["optional"] is True
            assert params[idx]["default"] is Empty
            assert params[idx]["mech"] == "ByVal"

    def test_optional_union_and_fallback_type_annotations(
        self: TestLibraryBuilder,
    ) -> None:
        project = (
            LibraryBuilder("TestLib").add_class("Cls", SampleClass).build()
        )
        params = project["classes"]["cls"]["functions"]["method_one"]["params"]

        assert params[0]["var"].declared_type == "integer"
        assert params[0]["optional"] is True
        assert params[0]["default"] is None

        assert params[1]["var"].declared_type == "string"
        assert params[1]["optional"] is True
        assert params[1]["default"] is None

        assert params[2]["var"].declared_type == "variant"
        assert params[2]["optional"] is True
        assert params[2]["default"] is Empty

    @pytest.mark.parametrize(
        ("annotation", "expected_type"),
        [
            (VBAInteger, "integer"),
            (VBALong, "long"),
            (VBAString, "string"),
            (VBANumericType, "variant"),
            (VBATypeBase, "variant"),
            (int, "variant"),
            (None, "variant"),
        ],
    )
    def test_make_vba_variable_type_mapping(
        self: TestLibraryBuilder, annotation: Any, expected_type: str
    ) -> None:
        var: VBAVariable = LibraryBuilder._make_vba_variable(annotation)
        assert isinstance(var, VBAVariable)
        assert var.declared_type == expected_type
