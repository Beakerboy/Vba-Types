from .vba_registry import registry
from vba_types.null import Null, VBANull


def handle_null_propogation(left: VBATypeBase, right: VBATypeBase) -> VBANull:
    return Null


registry.register("+", VBANull, VBATypeBase, handle_null_propogation)
registry.register("+", VBATypeBase, VBANull, handle_null_propogation)
