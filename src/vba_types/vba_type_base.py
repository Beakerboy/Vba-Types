from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Any


class VBATypeBase(ABC):
    def __init__(self, value: Any):
        self.value = value
