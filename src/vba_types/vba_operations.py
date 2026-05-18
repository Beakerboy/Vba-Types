from .vba_registry import registry
from vba_types.null import Null, VBANull
from vba_types.vba_type_base import VBATypeBase


def handle_null_propogation(left: VBATypeBase, right: VBATypeBase) -> VBANull:
    return Null


registry.register("+", VBANull, VBATypeBase, handle_null_propogation)
registry.register("+", VBATypeBase, VBANull, handle_null_propogation)
registry.register("-", VBANull, VBATypeBase, handle_null_propogation)
registry.register("-", VBATypeBase, VBANull, handle_null_propogation)
registry.register("*", VBANull, VBATypeBase, handle_null_propogation)
registry.register("*", VBATypeBase, VBANull, handle_null_propogation)
registry.register("/", VBANull, VBATypeBase, handle_null_propogation)
registry.register("/", VBATypeBase, VBANull, handle_null_propogation)
