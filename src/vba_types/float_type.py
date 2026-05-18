from __future__ import annotations
from .numeric_type import VBANumericType


T = TypeVar("T", bound="VBAFloatType")


class VBAFloatType(VBANumericType):
    pass
