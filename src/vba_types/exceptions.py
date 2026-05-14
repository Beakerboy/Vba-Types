from typing import Type, TypeVar, Dict, Optional, Union


T = TypeVar('T', bound='VBAMetadata')


class VBAMetadata:
    """
    Mixin for VBA metadata; avoids memory layout conflicts with C-based
    exceptions.
    """
    code: int
    message: str

    def __str__(self: T) -> str:
        return f"Run-time error '{self.code}': {self.message}"


class VBAException(VBAMetadata, Exception):
    """Base class for VBA errors with no close Python equivalent."""
    pass

# --- Native Python Mappings ---
# These inherit from standard Python exceptions for better interoperability.


class OverflowException(VBAMetadata, OverflowError):
    code = 6
    message = "Overflow"


class SubscriptOutOfRangeError(VBAMetadata, IndexError):
    code = 9
    message = "Subscript out of range"


class DivisionByZeroError(VBAMetadata, ZeroDivisionError):
    code = 11
    message = "Division by zero"


class TypeMismatchError(VBAMetadata, TypeError):
    code = 13
    message = "Type mismatch"


class OutOfStackSpaceError(VBAMetadata, RecursionError):
    code = 28
    message = "Out of stack space"


class FileNotFoundException(VBAMetadata, FileNotFoundError):
    code = 53
    message = "File not found"


class PermissionDeniedError(VBAMetadata, PermissionError):
    code = 70
    message = "Permission denied"


class PathNotFoundError(FileNotFoundError, VBAMetadata):
    code = 76
    message = "Path not found"


class ObjectDoesntSupportPropertyOrMethodError(VBAMetadata, AttributeError):
    code = 438
    message = "Object doesn't support this property or method"

# --- Unique VBA Exceptions ---
# These describe states unique to the VBA environment.


class ReturnWithoutGosubError(VBAException):
    code = 3
    message = "Return without GoSub"


class InvalidProcedureCallError(VBAException):
    code = 5
    message = "Invalid procedure call"


class OutOfMemoryError(VBAException):
    code = 7
    message = "Out of memory"


class ArrayLockedError(VBAException):
    code = 10
    message = "This array is fixed or temporarily locked"


class OutOfStringSpaceError(VBAException):
    code = 14
    message = "Out of string space"


class UserInterruptError(VBAException):
    code = 18
    message = "User interrupt occurred"


class ResumeWithoutError(VBAException):
    code = 20
    message = "Resume without error"


class ObjectVariableNotSetError(VBAException):
    code = 91
    message = "Object variable or With block variable not set"


class ActiveXComponentCantCreateObjectError(VBAException):
    code = 429
    message = "ActiveX component can't create object"


class AutomationError(VBAException):
    code = 440
    message = "Automation error"


class ApplicationDefinedError(VBAException):
    code = 1004
    message = "Application-defined or object-defined error"
