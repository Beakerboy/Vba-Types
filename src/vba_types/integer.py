from __future__ import annotations
from .integral_type import VBAIntegralType
from typing import TypeVar


T = TypeVar("T", bound="VBAInteger")


class VBAInteger(VBAIntegralType):
    """
    Simulates the VBA Integer data type (16-bit signed).
    Range: -32,768 to 32,767.
    """
    MIN_VALUE: int = -32768
    MAX_VALUE: int = 32767
