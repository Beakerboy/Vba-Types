from __future__ import annotations
from .vba_type_base import VBATypeBase
from typing import Any, TypeVar, TYPE_CHECKING


T = TypeVar("T", bound="VBANumericType")


class VBANumericType(VBATypeBase):
    value: int
