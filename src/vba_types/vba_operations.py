from .vba_registry import registry
from vba_types.null import Null, VBANull
from vba_types.vba_type_base import VBATypeBase
from vba_types.boolean import VBABoolean
from vba_types.empty import VBAEmpty



def handle_null_propogation(left: VBATypeBase, right: VBATypeBase) -> VBANull:
    return Null

def add_promote_to_integer(left: VBATypeBase, right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value + right.value)

def _subpromote_to_integer(left: VBATypeBase, right: VBATypeBase) -> VBAInteger:
    return VBAInteger(left.value + right.value)

registry.register("+", VBANull, VBATypeBase, handle_null_propogation)
registry.register("+", VBATypeBase, VBANull, handle_null_propogation)
registry.register("-", VBANull, VBATypeBase, handle_null_propogation)
registry.register("-", VBATypeBase, VBANull, handle_null_propogation)
registry.register("*", VBANull, VBATypeBase, handle_null_propogation)
registry.register("*", VBATypeBase, VBANull, handle_null_propogation)
registry.register("/", VBANull, VBATypeBase, handle_null_propogation)
registry.register("/", VBATypeBase, VBANull, handle_null_propogation)

registry.register("+", VBAInteger, VBAInteger, add_promote_to_integer)
registry.register("+", VBABoolean, VBAInteger, add_promote_to_integer)
registry.register("+", VBAInteger, VBABoolean, add_promote_to_integer)
registry.register("+", VBAEmpty, VBAEmpty, add_promote_to_integer)
registry.register("-", VBAInteger, VBAInteger, sub_promote_to_integer)
registry.register("-", VBABoolean, VBAInteger, sub_promote_to_integer)
registry.register("-", VBAInteger, VBABoolean, sub_promote_to_integer)
registry.register("-", VBAEmpty, VBAEmpty, sub_promote_to_integer)
