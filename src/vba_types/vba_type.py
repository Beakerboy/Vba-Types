from typing import TypeVar
from .boolean import VBABoolean


T = TypeVar('T', bound='VbaType')


class VbaType:
    def __eq__(self: T, other) -> bool:
        pass
